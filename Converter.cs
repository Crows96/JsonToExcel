using System.Buffers;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace JsonToExcel;

public enum PriceMode { Min, Max, Average, Median }

/// <summary>Как делить выгрузку по гео (target_regions / region / target_shops).</summary>
public enum RegionMode
{
    Ignore,       // регионы не учитываем
    AllRegions,   // один Excel: только товары, которые есть во всех регионах
    PerRegion,    // отдельный Excel (папка) на каждый регион (название папки = значение из target_regions)
    PerShop       // отдельный Excel (папка) на каждый магазин/адрес (название папки = значение из target_shops)
}

public sealed record ConvertOptions(
    string JsonSource,          // путь к файлу или ссылка
    string OutDir,
    decimal MaxSpreadPercent,   // допустимый разброс цен, %
    PriceMode PriceMode,
    int MaxPartMb,              // макс. размер картинок в одной части
    bool MakeZip,
    int MaxProducts,            // 0 — без ограничения (в каждом Excel)
    bool MergeSame,             // объединять товары с одинаковым названием и картинкой
    bool ShopFilter,            // фильтр по магазинам (target_shops)
    decimal ShopPercent,        // минимум % магазинов
    RegionMode RegionMode,
    RecodeOptions Recode,       // сжатие картинок
    int DownloadThreads,        // сколько картинок скачивать одновременно
    bool Economy);              // режим для слабого ПК: меньше потоков и нагрузки

public static class Converter
{
    // ======================= Модель данных =======================

    // Цена одного оффера (0 — цены нет) и код каталога (гео): >= 0 — один каталог,
    // -1 — нет каталога, <= -2 — оффер лежит в нескольких каталогах (индекс в GeoModel.Multi = -Cat - 2)
    private readonly record struct PriceEntry(decimal Price, int Cat);

    private sealed class Group
    {
        public string Description = "";
        public string? ImageUrl;
        public List<PriceEntry> Offers = new();
        public HashSet<int> Shops = new();      // уникальные магазины (target_shops), где есть товар
        public HashSet<int> Regions = new();    // регионы (target_regions / region), где есть товар
        public DateTime? DateStart;
        public DateTime? DateEnd;
    }

    private sealed class GeoModel
    {
        public List<string> RegionNames = new();
        public List<string> ShopNames = new();                            // индекс магазина -> значение из target_shops
        public List<int[]> Multi = new();                                 // офферы, лежащие в нескольких каталогах
        public int[][] CatRegions = Array.Empty<int[]>();                 // каталог -> регионы
        public int[][] CatShops = Array.Empty<int[]>();                   // каталог -> магазины
        public Dictionary<int, HashSet<int>> RegionShops = new();         // регион -> магазины его каталогов
        public HashSet<int> UsedShops = new();                            // магазины, в которых реально есть товары
        public HashSet<int> UsedRegions = new();                          // регионы, в которых реально есть товары
    }

    // Одна выгрузка: либо весь файл, либо один регион
    private sealed class Scope
    {
        public string Title = "";
        public string Dir = "";
        public int Region = -1;                         // -1 — все регионы
        public int Shop = -1;                           // -1 — не по магазину; иначе выгрузка одного магазина
        public string Kind = "Регион";                  // подпись для лога: «Регион» / «Магазин»
        public List<Group> Candidates = new();
        public HashSet<int> ShopSet = new();            // магазины, по которым считаем охват
    }

    private sealed class SelStats
    {
        public int Dropped, NoPrice, Single, Equal, Similar, ShopDropped;
    }

    private sealed record Row(string Description, string? ImageFile, decimal Price, DateTime? DateStart, DateTime? DateEnd);

    private sealed class Part
    {
        public List<Row> Rows = new();
        public HashSet<string> Images = new(StringComparer.OrdinalIgnoreCase);
        public long Bytes;
    }

    // ======================= Точка входа =======================

    public static bool IsUrl(string s) =>
        Uri.TryCreate(s?.Trim(), UriKind.Absolute, out var u) &&
        (u.Scheme == Uri.UriSchemeHttp || u.Scheme == Uri.UriSchemeHttps);

    /// <summary>Возвращает путь к папке с результатом.</summary>
    public static async Task<string> RunAsync(
        ConvertOptions opt, IProgress<string> log, IProgress<int> percent, CancellationToken ct)
    {
        var createdParts = new List<string>();   // папки частей, созданные в этом запуске
        var runner = new Runner(opt, log, percent, ct, createdParts);
        try
        {
            return await runner.RunAsync();
        }
        catch (OperationCanceledException)
        {
            // Пользователь нажал «Стоп»: удаляем всё временное и незаконченное
            CleanupAfterCancel(opt.OutDir, createdParts);
            log.Report("Остановлено. Временные файлы удалены.");
            throw;
        }
    }

    /// <summary>Обрабатывает несколько JSON по очереди; каждый файл — отдельная выгрузка в своей папке.</summary>
    public static async Task<string> RunBatchAsync(
        IReadOnlyList<ConvertOptions> items, IProgress<string> log, IProgress<int> percent, CancellationToken ct)
    {
        int n = items.Count;
        int ok = 0;
        var errors = new List<string>();

        try
        {
            for (int i = 0; i < n; i++)
            {
                ct.ThrowIfCancellationRequested();
                var item = items[i];
                if (n > 1) log.Report($"########## Файл {i + 1} из {n}: {item.JsonSource} ##########");

                int fileIndex = i;
                var filePercent = new DelegateProgress<int>(p => percent.Report((fileIndex * 100 + p) / n));
                try
                {
                    await RunAsync(item, log, filePercent, ct);
                    ok++;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex) when (n > 1)
                {
                    // Один «плохой» файл не должен останавливать остальные
                    errors.Add($"{item.JsonSource}: {ex.Message}");
                    log.Report($"ОШИБКА в файле {item.JsonSource}: {ex.Message}");
                }

                // Между файлами освобождаем память
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        catch (OperationCanceledException)
        {
            if (n > 1)
            {
                try
                {
                    var parent = Path.GetDirectoryName(Path.GetFullPath(items[0].OutDir));
                    if (parent != null && Directory.Exists(parent) && !Directory.EnumerateFileSystemEntries(parent).Any())
                        Directory.Delete(parent);
                }
                catch { }
            }
            throw;
        }

        percent.Report(100);
        var sb = new StringBuilder();
        sb.Append($"Обработано файлов: {ok} из {n}");
        foreach (var e in errors) sb.Append("\n  ✗ ").Append(e);
        return sb.ToString();
    }

    private static void CleanupAfterCancel(string outDir, List<string> createdParts)
    {
        string fullOut = Path.GetFullPath(outDir);

        foreach (var p in createdParts)
        {
            try { if (Directory.Exists(p)) Directory.Delete(p, true); } catch { }

            // Если часть лежала в папке региона и там больше ничего нет — убираем и папку региона
            try
            {
                var parent = Path.GetDirectoryName(Path.GetFullPath(p));
                if (parent != null && !string.Equals(parent, fullOut, StringComparison.OrdinalIgnoreCase) &&
                    Directory.Exists(parent) && !Directory.EnumerateFileSystemEntries(parent).Any())
                    Directory.Delete(parent);
            }
            catch { }
        }
        try
        {
            string cache = Path.Combine(outDir, "_cache");
            if (Directory.Exists(cache)) Directory.Delete(cache, true);
        }
        catch { }
        try
        {
            string src = Path.Combine(outDir, "_source.json");
            if (File.Exists(src)) File.Delete(src);
        }
        catch { }
        try
        {
            // Папку результата убираем, только если в ней ничего не осталось
            if (Directory.Exists(outDir) && !Directory.EnumerateFileSystemEntries(outDir).Any())
                Directory.Delete(outDir);
        }
        catch { }
    }

    // ======================= Основной процесс =======================

    private sealed class Runner
    {
        private readonly ConvertOptions opt;
        private readonly IProgress<string> log;
        private readonly IProgress<int> percent;
        private readonly CancellationToken ct;
        private readonly List<string> createdParts;

        private readonly string outDir;
        private readonly string cacheDir;                       // временная папка для скачанных картинок
        private readonly HttpClient http = new() { Timeout = TimeSpan.FromSeconds(60) };
        private readonly int downloadThreads;
        private readonly SemaphoreSlim encodeGate;                  // сколько картинок одновременно сжимается
        private long bytesIn, bytesOut;                             // объём картинок до и после сжатия

        // Имя картинки в Excel/частях (читаемое, уникальное) и имя файла в кэше (по хэшу ссылки)
        private readonly Dictionary<string, string> urlToName = new();
        private readonly Dictionary<string, string> nameToCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> usedNames = new(StringComparer.OrdinalIgnoreCase);
        private readonly ConcurrentBag<string> failed = new();

        private string? tempJson;
        private int scopeIndex;
        private int scopeTotal = 1;

        public Runner(ConvertOptions opt, IProgress<string> log, IProgress<int> percent,
                      CancellationToken ct, List<string> createdParts)
        {
            this.opt = opt;
            this.log = log;
            this.percent = percent;
            this.ct = ct;
            this.createdParts = createdParts;
            outDir = opt.OutDir;
            cacheDir = Path.Combine(outDir, "_cache");
            http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");

            downloadThreads = Math.Clamp(opt.DownloadThreads, 1, 64);
            int encodeThreads = opt.Economy ? 1 : Math.Max(1, Environment.ProcessorCount - 1);
            encodeGate = new SemaphoreSlim(encodeThreads, encodeThreads);
        }

        public async Task<string> RunAsync()
        {
            try { return await CoreAsync(); }
            finally { http.Dispose(); encodeGate.Dispose(); }
        }

        // Общий прогресс: каждая выгрузка занимает свою долю шкалы (внутри: 0–80% скачивание, 80–99% упаковка)
        private void Pct(int local) =>
            percent.Report(Math.Min(99, (scopeIndex * 100 + Math.Clamp(local, 0, 100)) / scopeTotal));

        private void Cleanup(bool keepCache)
        {
            if (tempJson != null)
            {
                try { File.Delete(tempJson); } catch { }
            }
            if (!keepCache)
            {
                try { Directory.Delete(cacheDir, true); }
                catch (Exception ex) { log.Report("Не удалось удалить временную папку _cache: " + ex.Message); }
            }
        }

        private async Task<string> CoreAsync()
        {
            Directory.CreateDirectory(outDir);
            Directory.CreateDirectory(cacheDir);

            // ---------- 0. Если указана ссылка — сначала скачиваем JSON ----------
            string jsonPath = opt.JsonSource.Trim();
            if (IsUrl(jsonPath))
            {
                tempJson = Path.Combine(outDir, "_source.json");
                await DownloadJsonAsync(jsonPath, tempJson, log, ct);
                jsonPath = tempJson;
            }

            // ---------- 1. Читаем JSON и группируем записи по товару ----------
            log.Report("Читаю JSON...");
            var (groups, offersCount, geo) = ReadGroups(jsonPath, ct);
            GC.Collect();   // служебные структуры чтения больше не нужны
            ct.ThrowIfCancellationRequested();
            log.Report($"Записей в JSON: {offersCount}, уникальных товаров (по коду): {groups.Count}");
            if (geo.UsedShops.Count > 0 || geo.UsedRegions.Count > 0)
                log.Report($"География: магазинов — {geo.UsedShops.Count}, регионов — {geo.UsedRegions.Count}");

            if (opt.MergeSame)
            {
                int before = groups.Count;
                groups = MergeSameProducts(groups);
                if (groups.Count < before)
                    log.Report($"Объединено товаров с одинаковым названием и картинкой: {before - groups.Count} " +
                               $"(осталось {groups.Count})");
            }

            // ---------- 2. Выгрузки: весь файл / один Excel на все регионы / Excel на каждый регион ----------
            var scopes = BuildScopes(groups, geo);
            if (scopes.Count == 0)
            {
                log.Report("Нет товаров для выгрузки.");
                Cleanup(false);
                percent.Report(100);
                return outDir;
            }

            scopeTotal = scopes.Count;
            for (scopeIndex = 0; scopeIndex < scopes.Count; scopeIndex++)
            {
                ct.ThrowIfCancellationRequested();
                await ProcessScopeAsync(scopes[scopeIndex]);
            }

            // Временную папку с картинками удаляем после упаковки. Если часть картинок не скачалась,
            // оставляем её: при повторном запуске докачаются только недостающие.
            Cleanup(keepCache: !failed.IsEmpty);

            if (opt.Recode.Enabled && bytesIn > 0)
                log.Report($"Картинки после сжатия: {bytesIn / 1048576.0:F0} МБ → {bytesOut / 1048576.0:F0} МБ " +
                           $"(в {(double)bytesIn / Math.Max(1, bytesOut):F1} раза меньше)");
            log.Report($"Готово. Результат: {outDir}");
            if (!failed.IsEmpty)
            {
                log.Report($"Не удалось скачать картинки ({failed.Count}), эти товары в Excel не добавлены:");
                foreach (var f in failed) log.Report("  " + f);
                log.Report("Папка _cache сохранена — запустите программу ещё раз, чтобы докачать недостающее.");
            }
            percent.Report(100);
            return outDir;
        }

        // ---------- Состав выгрузок ----------
        private List<Scope> BuildScopes(List<Group> groups, GeoModel geo)
        {
            var scopes = new List<Scope>();
            var mode = opt.RegionMode;

            if ((mode == RegionMode.AllRegions || mode == RegionMode.PerRegion) && geo.UsedRegions.Count == 0)
            {
                log.Report("В файле нет регионов (target_regions / region) — режим регионов не применяется.");
                mode = RegionMode.Ignore;
            }
            if (mode == RegionMode.PerShop && geo.UsedShops.Count == 0)
            {
                log.Report("В файле нет магазинов (target_shops) — разбивка по магазинам не применяется.");
                mode = RegionMode.Ignore;
            }

            switch (mode)
            {
                case RegionMode.AllRegions:
                {
                    int all = geo.UsedRegions.Count;
                    var cand = groups.Where(g => g.Regions.Count == all).ToList();
                    log.Report($"Режим регионов: один Excel — только товары, которые есть во всех регионах ({all}): {cand.Count}");
                    if (cand.Count == 0)
                    {
                        int best = groups.Count == 0 ? 0 : groups.Max(g => g.Regions.Count);
                        log.Report($"Ни один товар не встречается во всех регионах (максимальный охват — {best} из {all}). " +
                                   "Попробуйте режим «Отдельный Excel на каждый регион».");
                        return scopes;
                    }
                    scopes.Add(new Scope { Dir = outDir, Candidates = cand, ShopSet = geo.UsedShops });
                    break;
                }
                case RegionMode.PerRegion:
                {
                    var usedDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (int r in geo.UsedRegions.OrderBy(x => x))
                    {
                        var cand = groups.Where(g => g.Regions.Contains(r)).ToList();
                        if (cand.Count == 0) continue;

                        var shopSet = new HashSet<int>();
                        if (geo.RegionShops.TryGetValue(r, out var rs))
                            foreach (var sh in rs)
                                if (geo.UsedShops.Contains(sh)) shopSet.Add(sh);

                        string folder = UniqueName(SafeFolderName(geo.RegionNames[r]), usedDirs);
                        scopes.Add(new Scope
                        {
                            Title = geo.RegionNames[r],
                            Region = r,
                            Dir = Path.Combine(outDir, folder),
                            Candidates = cand,
                            ShopSet = shopSet
                        });
                    }
                    log.Report($"Режим регионов: отдельный Excel на каждый регион — регионов: {scopes.Count}");
                    break;
                }
                case RegionMode.PerShop:
                {
                    var usedDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (int s in geo.UsedShops.OrderBy(x => x))
                    {
                        var cand = groups.Where(g => g.Shops.Contains(s)).ToList();
                        if (cand.Count == 0) continue;

                        string shopName = geo.ShopNames[s];
                        string folder = UniqueName(SafeFolderName(shopName, "shop"), usedDirs);
                        scopes.Add(new Scope
                        {
                            Title = shopName,
                            Kind = "Магазин",
                            Shop = s,
                            Dir = Path.Combine(outDir, folder),
                            Candidates = cand,
                            ShopSet = new HashSet<int> { s }
                        });
                    }
                    log.Report($"Режим: отдельный Excel на каждый магазин (target_shops) — магазинов: {scopes.Count}");
                    break;
                }
                default:
                    scopes.Add(new Scope { Dir = outDir, Candidates = groups, ShopSet = geo.UsedShops });
                    break;
            }
            return scopes;
        }

        // ---------- Одна выгрузка: отбор, картинки, части ----------
        private async Task ProcessScopeAsync(Scope scope)
        {
            string tag = scope.Title.Length > 0 ? $"[{scope.Title}] " : "";
            void L(string msg) => log.Report(tag + msg);

            if (scope.Title.Length > 0) log.Report($"=== {scope.Kind}: {scope.Title} ===");

            // --- Отбор: охват по магазинам и разброс цен ---
            bool useShops = false;
            int needShops = 0;
            if (opt.ShopFilter && scope.Shop < 0)   // в выгрузке одного магазина охват по магазинам не нужен
            {
                if (scope.ShopSet.Count > 1)
                {
                    needShops = (int)Math.Ceiling(scope.ShopSet.Count * opt.ShopPercent / 100m);
                    useShops = needShops > 0;
                }
                else L("Магазины (target_shops) не заданы или магазин один — фильтр по магазинам не применён.");
            }

            var (selected, st) = SelectProducts(scope, useShops, needShops);

            if (useShops && selected.Count == 0)
            {
                int best = scope.Candidates.Count == 0 ? 0 : scope.Candidates.Max(g => CountIn(g.Shops, scope.ShopSet));
                L($"Ни один товар не встречается минимум в {opt.ShopPercent}% магазинов ({needShops} из {scope.ShopSet.Count}); " +
                  $"максимальный охват — {best}. Чтобы не получить пустой Excel, беру товары без фильтра по магазинам.");
                useShops = false;
                (selected, st) = SelectProducts(scope, false, 0);
            }

            if (useShops)
                L($"Фильтр по магазинам: товар должен быть минимум в {needShops} из {scope.ShopSet.Count} ({opt.ShopPercent}%); " +
                  $"отсеяно: {st.ShopDropped}");
            L($"Порог разброса цен: {opt.MaxSpreadPercent}%");
            L($"Взято товаров: {selected.Count} " +
              $"(цена одинаковая: {st.Equal}, разница в пределах порога: {st.Similar}, единственная цена: {st.Single})");
            L($"Пропущено из-за разброса цен: {st.Dropped}");
            if (st.NoPrice > 0) L($"Пропущено без корректной цены: {st.NoPrice}");

            if (selected.Count == 0)
            {
                L("Нет товаров для выгрузки.");
                Pct(100);
                return;
            }

            // --- Картинки (скачиваются в общий кэш, повторно не качаются) ---
            int noImageUrl = 0, imageFailed = 0;

            if (opt.MaxProducts > 0)
            {
                // Лимит (на каждый Excel): берём товары по порядку и качаем картинки партиями,
                // пока не наберём нужное число. Товары без картинки заменяются следующими.
                L($"Лимит: не больше {opt.MaxProducts} товаров");
                var accepted = new List<(Group Group, decimal Price)>();
                int idx = 0;
                while (accepted.Count < opt.MaxProducts && idx < selected.Count)
                {
                    ct.ThrowIfCancellationRequested();
                    int need = opt.MaxProducts - accepted.Count;
                    var batch = selected.GetRange(idx, Math.Min(need, selected.Count - idx));
                    idx += batch.Count;

                    await EnsureImagesAsync(batch.Select(b => b.Group.ImageUrl), null);

                    foreach (var item in batch)
                    {
                        string? url = item.Group.ImageUrl;
                        if (string.IsNullOrWhiteSpace(url)) { noImageUrl++; continue; }
                        if (!HasImage(url)) { imageFailed++; continue; }
                        accepted.Add(item);
                    }
                    Pct(accepted.Count * 80 / opt.MaxProducts);
                    L($"Подготовлено товаров: {accepted.Count}/{opt.MaxProducts}");
                }

                if (accepted.Count < opt.MaxProducts)
                    L($"Подходящих товаров меньше лимита: взято {accepted.Count} из {opt.MaxProducts}.");
                selected = accepted;
            }
            else
            {
                await EnsureImagesAsync(selected.Select(s => s.Group.ImageUrl), (d, t) =>
                {
                    Pct(d * 80 / t);                              // 0–80% — скачивание
                    if (d % 200 == 0 || d == t) L($"Скачано: {d}/{t}");
                });
            }

            // --- Делим товары на части по размеру картинок ---
            long maxBytes = (long)opt.MaxPartMb * 1024 * 1024;
            var parts = new List<Part>();
            var cur = new Part();
            foreach (var (g, price) in selected)
            {
                // Нет картинки (ссылки нет или скачать не удалось) — товар в Excel не добавляем
                string? url = g.ImageUrl;
                if (string.IsNullOrWhiteSpace(url)) { noImageUrl++; continue; }
                if (!HasImage(url)) { imageFailed++; continue; }

                string file = urlToName[url];
                long size = new FileInfo(CachePath(file)).Length;

                // Картинка, которая уже есть в текущей части, место не занимает
                long add = cur.Images.Contains(file) ? 0 : size;
                if (cur.Rows.Count > 0 && cur.Bytes + add > maxBytes)
                {
                    parts.Add(cur);
                    cur = new Part();
                }

                cur.Rows.Add(new Row(g.Description, file, price, g.DateStart, g.DateEnd));
                if (cur.Images.Add(file)) cur.Bytes += size;
            }
            if (cur.Rows.Count > 0) parts.Add(cur);

            if (noImageUrl + imageFailed > 0)
                L($"Не добавлено из-за картинок: {noImageUrl + imageFailed} " +
                  $"(нет ссылки на картинку: {noImageUrl}, не удалось скачать: {imageFailed})");
            if (parts.Count == 0)
            {
                L("Ни один товар не попал в Excel: ни для одного не удалось получить картинку.");
                Pct(100);
                return;
            }
            L($"Частей: {parts.Count} (не больше {opt.MaxPartMb} МБ картинок в каждой)");

            // Удаляем части от предыдущего запуска в этой папке, чтобы не оставалось лишнего
            if (Directory.Exists(scope.Dir))
                foreach (var old in Directory.GetDirectories(scope.Dir, "part_*"))
                    Directory.Delete(old, true);
            Directory.CreateDirectory(scope.Dir);

            // --- Пишем части: Excel + картинки (папка или ZIP) ---
            for (int i = 0; i < parts.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var part = parts[i];
                string partDir = Path.Combine(scope.Dir, $"part_{i + 1:D3}");
                Directory.CreateDirectory(partDir);
                createdParts.Add(partDir);

                WriteExcel(part, Path.Combine(partDir, "catalog.xlsx"));

                if (opt.MakeZip)
                {
                    string zipPath = Path.Combine(partDir, "images.zip");
                    using var zs = new FileStream(zipPath, FileMode.Create);
                    using var za = new ZipArchive(zs, ZipArchiveMode.Create);
                    foreach (var img in part.Images)
                    {
                        ct.ThrowIfCancellationRequested();
                        za.CreateEntryFromFile(CachePath(img), img, CompressionLevel.NoCompression);   // картинки уже сжаты — архив только упаковывает
                    }
                }
                else
                {
                    string imgDir = Path.Combine(partDir, "images");
                    Directory.CreateDirectory(imgDir);
                    foreach (var img in part.Images)
                    {
                        ct.ThrowIfCancellationRequested();
                        File.Copy(CachePath(img), Path.Combine(imgDir, img), true);
                    }
                }

                L($"part_{i + 1:D3}: товаров {part.Rows.Count}, картинок {part.Images.Count}, " +
                  $"{part.Bytes / 1024.0 / 1024.0:F1} МБ");
                Pct(80 + (i + 1) * 19 / parts.Count);             // 80–99% — упаковка
            }
        }

        private static void WriteExcel(Part part, string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Лист1");
            string[] headers = { "description", "image", "price_new", "date_start", "date_end" };
            for (int c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
            ws.Row(1).Style.Font.Bold = true;

            int row = 2;
            foreach (var r in part.Rows)
            {
                ws.Cell(row, 1).Value = r.Description;
                if (r.ImageFile != null) ws.Cell(row, 2).Value = r.ImageFile;
                ws.Cell(row, 3).Value = r.Price;
                // Даты берём из JSON, если они там есть; иначе ячейки остаются пустыми
                if (r.DateStart.HasValue) ws.Cell(row, 4).Value = r.DateStart.Value;
                if (r.DateEnd.HasValue) ws.Cell(row, 5).Value = r.DateEnd.Value;
                row++;
            }
            ws.Column(4).Style.DateFormat.Format = "dd.mm.yyyy";
            ws.Column(5).Style.DateFormat.Format = "dd.mm.yyyy";
            // AdjustToContents на больших листах очень медленный — задаём ширину колонок явно
            ws.Column(1).Width = 60;
            ws.Column(2).Width = 40;
            ws.Column(3).Width = 12;
            ws.Column(4).Width = 12;
            ws.Column(5).Width = 12;
            wb.SaveAs(path);
        }

        // ---------- Отбор товаров внутри одной выгрузки ----------
        private (List<(Group Group, decimal Price)> Items, SelStats Stats) SelectProducts(
            Scope scope, bool useShops, int needShops)
        {
            decimal limit = opt.MaxSpreadPercent / 100m;
            var st = new SelStats();
            var list = new List<(Group Group, decimal Price)>();

            foreach (var g in scope.Candidates)
            {
                if (useShops && CountIn(g.Shops, scope.ShopSet) < needShops) { st.ShopDropped++; continue; }

                var prices = ScopePrices(g, scope);
                if (prices.Count == 0) { st.NoPrice++; continue; }

                decimal min = prices.Min();
                decimal max = prices.Max();
                decimal spread = (max - min) / min;          // min > 0 гарантировано при чтении

                if (spread > limit) { st.Dropped++; continue; }

                if (prices.Count == 1) st.Single++;
                else if (max == min) st.Equal++;
                else st.Similar++;

                list.Add((g, PickPrice(prices, opt.PriceMode)));
            }
            return (list, st);
        }

        // Для выгрузки по региону цены берём только из офферов этого региона,
        // чтобы цены других регионов не влияли на отбор и итоговую цену.
        private List<decimal> ScopePrices(Group g, Scope scope)
        {
            var result = new List<decimal>(g.Offers.Count);
            foreach (var e in g.Offers)
            {
                if (e.Price <= 0) continue;                      // оффер без корректной цены
                if (scope.Region >= 0 && !InRegion(e, scope.Region)) continue;
                if (scope.Shop >= 0 && !InShop(e, scope.Shop)) continue;
                result.Add(e.Price);
            }
            return result;
        }

        private GeoModel? currentGeo;   // заполняется при чтении, нужен для проверки региона оффера

        private bool InRegion(PriceEntry e, int region) =>
            currentGeo != null && CodeHas(currentGeo, currentGeo.CatRegions, e.Cat, region);

        private bool InShop(PriceEntry e, int shop) =>
            currentGeo != null && CodeHas(currentGeo, currentGeo.CatShops, e.Cat, shop);

        private static bool CodeHas(GeoModel geo, int[][] sets, int code, int value)
        {
            if (code >= 0) return Array.IndexOf(sets[code], value) >= 0;
            if (code <= -2)
                foreach (var ci in geo.Multi[-code - 2])
                    if (Array.IndexOf(sets[ci], value) >= 0) return true;
            return false;
        }

        // Добавляет магазины и регионы каталога (или нескольких каталогов) в товар
        private static void AddGeo(Group g, GeoModel geo, int code)
        {
            if (code >= 0)
            {
                g.Shops.UnionWith(geo.CatShops[code]);
                g.Regions.UnionWith(geo.CatRegions[code]);
            }
            else if (code <= -2)
            {
                foreach (var ci in geo.Multi[-code - 2])
                {
                    g.Shops.UnionWith(geo.CatShops[ci]);
                    g.Regions.UnionWith(geo.CatRegions[ci]);
                }
            }
        }

        // ---------- Картинки ----------
        private string CachePath(string displayName) => Path.Combine(cacheDir, nameToCache[displayName]);

        // Картинка считается готовой, если файл скачан и не пустой
        private bool HasImage(string url)
        {
            if (!urlToName.TryGetValue(url, out var n)) return false;
            var fi = new FileInfo(CachePath(n));
            return fi.Exists && fi.Length > 0;
        }

        // Назначает имена файлов и скачивает (и при необходимости сжимает) картинки, которые ещё не обрабатывались
        private async Task EnsureImagesAsync(IEnumerable<string?> urls, Action<int, int>? onProgress)
        {
            var todo = new Dictionary<string, string>();
            foreach (var url in urls.Where(u => !string.IsNullOrWhiteSpace(u)).Distinct())
            {
                if (urlToName.ContainsKey(url!)) continue;      // уже обработана раньше

                string name = MakeFileName(url!);
                // При сжатии формат меняется, поэтому и расширение в имени (и в Excel) берём новое
                if (opt.Recode.Enabled) name = Path.GetFileNameWithoutExtension(name) + opt.Recode.Extension;
                string unique = name;
                int k = 1;
                while (!usedNames.Add(unique))
                    unique = $"{Path.GetFileNameWithoutExtension(name)}_{k++}{Path.GetExtension(name)}";

                urlToName[url!] = unique;
                nameToCache[unique] = CacheName(url!, unique, opt.Recode.Signature);
                todo[url!] = unique;
            }
            if (todo.Count == 0) return;
            if (onProgress != null) log.Report($"Картинок к скачиванию: {todo.Count}");

            int done = 0;
            int total = todo.Count;
            await Parallel.ForEachAsync(todo,
                new ParallelOptions { MaxDegreeOfParallelism = downloadThreads, CancellationToken = ct },
                async (pair, token) =>
                {
                    string path = CachePath(pair.Value);
                    if (!(File.Exists(path) && new FileInfo(path).Length > 0))   // уже скачано — пропускаем
                    {
                        for (int attempt = 1; attempt <= 3; attempt++)
                        {
                            try
                            {
                                await FetchImageAsync(pair.Key, path, token);
                                break;
                            }
                            catch (OperationCanceledException) when (token.IsCancellationRequested)
                            {
                                throw;                       // «Стоп» — не считаем ошибкой скачивания
                            }
                            catch (InvalidDataException ex)
                            {
                                failed.Add($"{pair.Key} -> {ex.Message}");   // не картинка — повторять бессмысленно
                                break;
                            }
                            catch (Exception) when (attempt < 3)
                            {
                                await Task.Delay(500 * attempt, token);
                            }
                            catch (Exception ex)
                            {
                                failed.Add($"{pair.Key} -> {ex.Message}");
                            }
                        }
                    }
                    int d = Interlocked.Increment(ref done);
                    onProgress?.Invoke(d, total);
                });
        }

        // Скачивает одну картинку в кэш. Пишет во временный файл и переименовывает после успеха,
        // поэтому недокачанный файл никогда не принимается за готовый.
        private async Task FetchImageAsync(string url, string path, CancellationToken token)
        {
            string tmp = path + ".part";
            try
            {
                using var resp = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, token);
                resp.EnsureSuccessStatusCode();

                if (!opt.Recode.Enabled)
                {
                    // Без сжатия: поток прямо на диск, картинка целиком в памяти не держится
                    await using (var src = await resp.Content.ReadAsStreamAsync(token))
                    await using (var dst = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
                        await src.CopyToAsync(dst, 81920, token);
                }
                else
                {
                    byte[] bytes = await resp.Content.ReadAsByteArrayAsync(token);

                    // Сжатие — самая тяжёлая для процессора часть, одновременно их идёт ограниченное число
                    byte[] result;
                    await encodeGate.WaitAsync(token);
                    try { result = ImageRecoder.Recode(bytes, opt.Recode).Data; }
                    finally { encodeGate.Release(); }

                    await File.WriteAllBytesAsync(tmp, result, token);
                    Interlocked.Add(ref bytesIn, bytes.Length);
                    Interlocked.Add(ref bytesOut, result.Length);
                }

                File.Move(tmp, path, overwrite: true);
            }
            catch
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                throw;
            }
        }

        // ---------- Чтение JSON ----------
        // JSON читается потоком (Utf8JsonReader): в памяти одновременно только один оффер, поэтому расход
        // памяти зависит от числа уникальных товаров, а не от размера файла. Два прохода по файлу:
        // сначала каталоги (география), затем офферы.
        private (List<Group> Groups, int OffersCount, GeoModel Geo) ReadGroups(string jsonPath, CancellationToken ct)
        {
            var list = new List<Group>();
            var index = new Dictionary<string, Group>();
            int count = 0;

            // --- География из каталогов: target_shops (магазины), target_regions или region (регионы) ---
            var geo = new GeoModel();
            var shopIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var regionIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var catRegions = new List<int[]>();
            var catShops = new List<int[]>();

            // 64-битный хэш id оффера -> каталог (или код нескольких каталогов, см. PriceEntry.Cat).
            // Хэш вместо строки-ключа: в разы меньше памяти на миллионах офферов.
            var offerCats = new Dictionary<ulong, int>();

            // Проход 1: каталоги
            ScanArray(jsonPath, "catalogs", c =>
            {
                int ci = catRegions.Count;

                var regions = new HashSet<int>();
                foreach (var name in GetStringList(c, "target_regions"))
                    regions.Add(Intern(regionIndex, geo.RegionNames, name));
                if (regions.Count == 0)
                {
                    string? one = GetString(c, "region");
                    if (!string.IsNullOrWhiteSpace(one)) regions.Add(Intern(regionIndex, geo.RegionNames, one));
                }

                var shops = new HashSet<int>();
                foreach (var name in GetStringList(c, "target_shops"))
                    shops.Add(Intern(shopIndex, geo.ShopNames, name));

                catRegions.Add(regions.ToArray());
                catShops.Add(shops.ToArray());

                foreach (var r in regions)
                {
                    if (!geo.RegionShops.TryGetValue(r, out var rs)) geo.RegionShops[r] = rs = new HashSet<int>();
                    rs.UnionWith(shops);
                }

                if (c.TryGetProperty("offers", out var ids) && ids.ValueKind == JsonValueKind.Array)
                {
                    foreach (var idEl in ids.EnumerateArray())
                    {
                        if (idEl.ValueKind != JsonValueKind.String) continue;
                        string? oid = idEl.GetString();
                        if (string.IsNullOrEmpty(oid)) continue;

                        ulong h = HashId(oid);
                        if (!offerCats.TryGetValue(h, out var code))
                        {
                            offerCats[h] = ci;
                        }
                        else if (code == ci)
                        {
                            // повтор id внутри одного каталога
                        }
                        else if (code >= 0)
                        {
                            geo.Multi.Add(new[] { code, ci });
                            offerCats[h] = -(geo.Multi.Count - 1) - 2;
                        }
                        else
                        {
                            int mi = -code - 2;
                            var old = geo.Multi[mi];
                            if (Array.IndexOf(old, ci) < 0)
                            {
                                var arr = new int[old.Length + 1];
                                old.CopyTo(arr, 0);
                                arr[^1] = ci;
                                geo.Multi[mi] = arr;
                            }
                        }
                    }
                }
            }, ct);

            geo.CatRegions = catRegions.ToArray();
            geo.CatShops = catShops.ToArray();
            currentGeo = geo;

            // Проход 2: офферы
            int nextLog = 10;
            bool found = ScanArray(jsonPath, "offers", o =>
            {
                count++;
                string id = GetString(o, "id") ?? "";

                // id вида "<код товара>_<регион/магазин>" -> ключ товара = часть до первого "_".
                // Если "_" нет (как в palich), каждая запись — отдельный товар.
                string key;
                int us = id.IndexOf('_');
                if (us > 0) key = id[..us];
                else key = id.Length > 0 ? id : (GetString(o, "description") ?? "");

                if (!index.TryGetValue(key, out var grp))
                {
                    grp = new Group();
                    index[key] = grp;
                    list.Add(grp);
                }

                // Текст, ссылку и даты читаем только пока они не найдены — лишних строк не создаём
                if (grp.Description.Length == 0) grp.Description = GetString(o, "description") ?? "";
                if (string.IsNullOrWhiteSpace(grp.ImageUrl)) grp.ImageUrl = GetString(o, "image");
                if (!grp.DateStart.HasValue) grp.DateStart = GetDate(o, "date_start");
                if (!grp.DateEnd.HasValue) grp.DateEnd = GetDate(o, "date_end");

                int cat = (id.Length > 0 && offerCats.TryGetValue(HashId(id), out var cc)) ? cc : -1;
                var price = GetDecimal(o, "price_new");
                grp.Offers.Add(new PriceEntry(price.HasValue && price.Value > 0 ? price.Value : 0m, cat));
            }, ct, p =>
            {
                if (p >= nextLog)
                {
                    log.Report($"Чтение товаров: {p}%");
                    nextLog = p / 10 * 10 + 10;
                }
            });

            if (!found) throw new Exception("В JSON нет массива \"offers\".");

            // Магазины и регионы товара считаем один раз по уникальным каталогам его офферов.
            // Учитываем только магазины и регионы, в которых реально есть товары (пустые каталоги не считаем).
            foreach (var g in list)
            {
                ct.ThrowIfCancellationRequested();
                var seen = new HashSet<int>();
                foreach (var e in g.Offers)
                    if (e.Cat != -1 && seen.Add(e.Cat)) AddGeo(g, geo, e.Cat);

                g.Offers.TrimExcess();
                geo.UsedShops.UnionWith(g.Shops);
                geo.UsedRegions.UnionWith(g.Regions);
            }

            return (list, count, geo);
        }
    }

    // ======================= Вспомогательные функции =======================

    // Потоковое чтение: находит в корне JSON массив с именем property и вызывает onElement для каждого
    // элемента-объекта. В памяти одновременно лежат только буфер (~1 МБ) и один элемент.
    // Возвращает false, если такого массива в файле нет.
    private static bool ScanArray(string path, string property, Action<JsonElement> onElement,
                                  CancellationToken ct, Action<int>? onPercent = null)
    {
        var options = new JsonReaderOptions
        {
            AllowTrailingCommas = true,
            CommentHandling = JsonCommentHandling.Skip
        };

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16,
                                      FileOptions.SequentialScan);
        long length = Math.Max(1, fs.Length);

        byte[] buffer = ArrayPool<byte>.Shared.Rent(1 << 20);
        try
        {
            int have = 0;                       // сколько байт сейчас лежит в буфере
            bool eof = false, first = true;
            bool inTarget = false, inArray = false, found = false, done = false;
            int counter = 0, lastPercent = -1;
            var state = new JsonReaderState(options);

            while (!done)
            {
                // Дочитываем буфер
                while (!eof && have < buffer.Length)
                {
                    int n = fs.Read(buffer, have, buffer.Length - have);
                    if (n == 0) { eof = true; break; }
                    have += n;
                }

                if (first)
                {
                    first = false;
                    // Utf8JsonReader сам BOM не пропускает
                    if (have >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
                    {
                        buffer.AsSpan(3, have - 3).CopyTo(buffer);
                        have -= 3;
                    }
                }

                var reader = new Utf8JsonReader(buffer.AsSpan(0, have), eof, state);
                bool needMore = false;

                while (true)
                {
                    var before = reader;
                    if (!reader.Read()) break;

                    var token = reader.TokenType;
                    int depth = reader.CurrentDepth;

                    if (depth == 1 && token == JsonTokenType.PropertyName)
                    {
                        inTarget = reader.ValueTextEquals(property);
                    }
                    else if (depth == 1 && token == JsonTokenType.StartArray && inTarget)
                    {
                        inArray = true;
                        found = true;
                    }
                    else if (depth == 1 && token == JsonTokenType.EndArray && inArray)
                    {
                        done = true;                    // нужный массив прочитан целиком
                        break;
                    }
                    else if (depth == 2 && token == JsonTokenType.StartObject && inArray)
                    {
                        // Элемент должен целиком поместиться в буфер, иначе дочитываем файл и повторяем
                        var probe = reader;
                        if (!probe.TrySkip())
                        {
                            reader = before;
                            needMore = true;
                            break;
                        }

                        using var doc = JsonDocument.ParseValue(ref reader);
                        onElement(doc.RootElement);

                        if ((++counter & 255) == 0)
                        {
                            ct.ThrowIfCancellationRequested();
                            if (onPercent != null)
                            {
                                int p = (int)(fs.Position * 100 / length);
                                if (p != lastPercent) { lastPercent = p; onPercent(p); }
                            }
                        }
                    }
                }

                if (done) break;
                if (eof && !needMore) break;            // файл закончился

                state = reader.CurrentState;
                int consumed = (int)reader.BytesConsumed;
                int left = have - consumed;
                if (left > 0 && consumed > 0) buffer.AsSpan(consumed, left).CopyTo(buffer);
                have = left;

                if (have == buffer.Length)              // один элемент не поместился — увеличиваем буфер
                {
                    var bigger = ArrayPool<byte>.Shared.Rent(buffer.Length * 2);
                    buffer.AsSpan(0, have).CopyTo(bigger);
                    ArrayPool<byte>.Shared.Return(buffer);
                    buffer = bigger;
                }
                ct.ThrowIfCancellationRequested();
            }
            return found;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    // 64-битный хэш (FNV-1a) строки: ключ для таблицы «оффер -> каталог» без хранения самих строк
    private static ulong HashId(string s)
    {
        ulong h = 14695981039346656037UL;
        foreach (char ch in s)
        {
            h ^= ch;
            h *= 1099511628211UL;
        }
        return h;
    }

    private static async Task DownloadJsonAsync(string url, string path, IProgress<string> log, CancellationToken ct)
    {
        log.Report("Скачиваю JSON по ссылке...");
        var handler = new HttpClientHandler { AutomaticDecompression = System.Net.DecompressionMethods.All };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromMinutes(15) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");

        try
        {
            using var resp = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
            resp.EnsureSuccessStatusCode();

            using var src = await resp.Content.ReadAsStreamAsync(ct);
            using var dst = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);

            var buf = new byte[81920];
            long total = 0, lastLogged = 0;
            int read;
            while ((read = await src.ReadAsync(buf.AsMemory(), ct)) > 0)
            {
                await dst.WriteAsync(buf.AsMemory(0, read), ct);
                total += read;
                if (total - lastLogged >= 10L * 1024 * 1024)
                {
                    lastLogged = total;
                    log.Report($"JSON: скачано {total / 1024.0 / 1024.0:F0} МБ...");
                }
            }
            log.Report($"JSON скачан: {total / 1024.0 / 1024.0:F1} МБ");
        }
        catch (Exception ex) when (!ct.IsCancellationRequested && ex is HttpRequestException or TaskCanceledException or IOException)
        {
            throw new Exception("Не удалось скачать JSON по ссылке: " + ex.Message);
        }
    }

    // Товары с разными кодами, но одинаковым названием И одинаковой ссылкой на картинку,
    // считаем одним товаром: их цены сравниваются вместе.
    // Если у товара нет названия или картинки, по этому признаку он не объединяется.
    private static List<Group> MergeSameProducts(List<Group> groups)
    {
        var result = new List<Group>();
        var index = new Dictionary<string, Group>();

        foreach (var g in groups)
        {
            string desc = NormalizeText(g.Description);
            string img = g.ImageUrl?.Trim() ?? "";

            if (desc.Length == 0 || img.Length == 0)
            {
                result.Add(g);
                continue;
            }

            string key = desc + "\u0001" + img;
            if (index.TryGetValue(key, out var first))
            {
                first.Offers.AddRange(g.Offers);
                first.Shops.UnionWith(g.Shops);
                first.Regions.UnionWith(g.Regions);
                first.DateStart ??= g.DateStart;
                first.DateEnd ??= g.DateEnd;
            }
            else
            {
                index[key] = g;
                result.Add(g);
            }
        }
        return result;
    }

    // Без учёта регистра и лишних пробелов
    private static string NormalizeText(string s) =>
        string.Join(' ', s.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant();

    private static int CountIn(HashSet<int> items, HashSet<int> set)
    {
        int n = 0;
        foreach (var x in items)
            if (set.Contains(x)) n++;
        return n;
    }

    private static int Intern(Dictionary<string, int> index, List<string>? names, string value)
    {
        value = value.Trim();
        if (!index.TryGetValue(value, out var i))
        {
            i = index.Count;
            index[value] = i;
            names?.Add(value);
        }
        return i;
    }

    private static IEnumerable<string> GetStringList(JsonElement e, string key)
    {
        if (!e.TryGetProperty(key, out var v) || v.ValueKind != JsonValueKind.Array) yield break;
        foreach (var x in v.EnumerateArray())
        {
            if (x.ValueKind != JsonValueKind.String) continue;
            var s = x.GetString();
            if (!string.IsNullOrWhiteSpace(s)) yield return s;
        }
    }

    private static string SafeFolderName(string name, string fallback = "region")
    {
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        name = Regex.Replace(name, @"\s+", " ").Trim().TrimEnd('.');
        if (name.Length > 80) name = name[..80].TrimEnd();
        return name.Length == 0 ? fallback : name;
    }

    private static string UniqueName(string name, HashSet<string> used)
    {
        string unique = name;
        int k = 2;
        while (!used.Add(unique)) unique = $"{name} ({k++})";
        return unique;
    }

    private static decimal PickPrice(List<decimal> prices, PriceMode mode)
    {
        switch (mode)
        {
            case PriceMode.Max:
                return prices.Max();
            case PriceMode.Average:
                return Math.Round(prices.Average(), 2);
            case PriceMode.Median:
            {
                var s = prices.OrderBy(p => p).ToList();
                int n = s.Count;
                return n % 2 == 1 ? s[n / 2] : Math.Round((s[n / 2 - 1] + s[n / 2]) / 2m, 2);
            }
            default:
                return prices.Min();
        }
    }

    private static string? GetString(JsonElement e, string key) =>
        e.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static DateTime? GetDate(JsonElement e, string key)
    {
        var s = GetString(e, key);
        // DateTimeOffset сохраняет календарную дату, записанную в JSON, даже если указан часовой пояс
        return DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
            ? d.Date : null;
    }

    private static decimal? GetDecimal(JsonElement e, string key)
    {
        if (!e.TryGetProperty(key, out var v)) return null;
        if (v.ValueKind == JsonValueKind.Number) return v.TryGetDecimal(out var n) ? n : null;
        if (v.ValueKind == JsonValueKind.String &&
            decimal.TryParse(v.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
        return null;
    }

    // Имя файла картинки для Excel и папки: из последней части ссылки.
    // Если имя типа «800x800.jpeg» (одинаковое у всех товаров), добавляем предыдущую часть пути.
    private static string MakeFileName(string url)
    {
        string name = "";
        try
        {
            var uri = new Uri(url);
            var segs = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (segs.Length > 0)
            {
                string last = segs[^1];
                if (segs.Length > 1 && Regex.IsMatch(Path.GetFileNameWithoutExtension(last), @"^\d{2,4}x\d{2,4}$"))
                    last = segs[^2] + "_" + last;
                name = Uri.UnescapeDataString(last);
            }
        }
        catch { name = ""; }

        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        name = name.Trim().TrimEnd('.');

        string ext = Path.GetExtension(name);
        string stem = Path.GetFileNameWithoutExtension(name);
        if (stem.Length > 90) stem = stem[..90];
        if (stem.Length == 0) stem = HashHex(url, 12);
        if (ext.Length == 0 || ext.Length > 6) ext = ".jpg";
        return stem + ext;
    }

    // Имя файла в кэше зависит от ссылки и настроек сжатия — устаревшие файлы от другого запуска
    // (с другим качеством или размером) не подменят картинку
    private static string CacheName(string url, string displayName, string signature) =>
        HashHex(url + "|" + signature, 20) + Path.GetExtension(displayName);

    private static string HashHex(string text, int length) =>
        Convert.ToHexString(SHA1.HashData(Encoding.UTF8.GetBytes(text)))[..length];
}

/// <summary>Синхронный IProgress: вызывает делегат сразу в том же потоке (Progress&lt;T&gt; может менять порядок вызовов).</summary>
internal sealed class DelegateProgress<T> : IProgress<T>
{
    private readonly Action<T> action;
    public DelegateProgress(Action<T> action) => this.action = action;
    public void Report(T value) => action(value);
}

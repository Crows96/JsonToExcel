using System.Buffers;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace JsonToExcel;

public enum PriceMode { Min, Max, Average, Median }

/// <summary>Формат исходного файла (определяется по содержимому).</summary>
public enum SourceFormat { Json, Xml }

/// <summary>Как делить выгрузку по гео (target_regions / region / target_shops).</summary>
public enum RegionMode
{
    Ignore,       // регионы не учитываем
    AllRegions,   // один Excel: только товары, которые есть во всех регионах
    PerRegion,    // отдельный Excel (папка) на каждый регион (название папки = значение из target_regions)
    PerShop,      // отдельный Excel (папка) на каждый магазин/адрес (название папки = значение из target_shops)
    RegionsMerged // один Excel: охват и цены считаются в каждом регионе отдельно; товар берём, если прошёл фильтр в каждом регионе, где он есть
}

/// <summary>Один файл (или ссылка) в режиме объединения. YML — отдельный магазин; в JSON магазины берутся из каталогов (target_shops).</summary>
public sealed record FeedSource(string Source, string Label, string? Region = null);

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
    bool Economy,               // режим для слабого ПК: меньше потоков и нагрузки
    bool PromoOnly,             // только акционные позиции: в оффере есть price_old (старая цена)
    IReadOnlyList<FeedSource>? Feeds = null,   // не null — режим объединения: все файлы списка (YML и/или JSON) в один результат
    bool RegionBySubject = false);             // регион = субъект РФ (область, край, республика) по справочнику, а не город

public static class Converter
{
    // YML-фиды часто в windows-1251: в .NET такие кодировки доступны только после регистрации провайдера
    static Converter() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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
        public decimal? OldPrice;               // старая (обычная) цена — есть в YML (oldprice)
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
        public List<(Group Group, decimal Price)>? Preselected;   // отбор уже сделан (режим «по регионам, один Excel»)
    }

    private sealed class SelStats
    {
        public int Dropped, NoPrice, Single, Equal, Similar, ShopDropped;
    }

    private sealed record Row(string Description, string? ImageFile, decimal Price, DateTime? DateStart, DateTime? DateEnd, decimal? OldPrice = null);

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
        var keptCache = new List<string>();      // файлы, у которых остался _cache (не все картинки скачались)

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
                    if (Directory.Exists(Path.Combine(item.OutDir, "_cache"))) keptCache.Add(item.JsonSource);
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
        if (n == 1 && items[0].Feeds != null)
        {
            // Режим объединения: сколько файлов списка вошло в результат
            int total = items[0].Feeds!.Count;
            string failedPath = Path.Combine(items[0].OutDir, "failed_feeds.txt");
            int failed = File.Exists(failedPath) ? File.ReadAllLines(failedPath).Length : 0;
            sb.Append($"Объединено файлов: {total - failed} из {total}");
            if (failed > 0)
                sb.Append($"\n⚠ Не прочитано: {failed}. Список для повторной загрузки: {failedPath}");
        }
        else
            sb.Append($"Обработано файлов: {ok} из {n}");
        foreach (var e in errors) sb.Append("\n  ✗ ").Append(e);
        if (keptCache.Count > 0)
        {
            sb.Append("\n⚠ Часть картинок не скачалась, папка _cache оставлена (повторный запуск докачает недостающее):");
            foreach (var k in keptCache) sb.Append("\n  ").Append(k);
        }
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
        private readonly ConcurrentBag<string> failed = new();        // временные сбои: повторный запуск может помочь
        private readonly ConcurrentBag<string> missing = new();       // картинки нет на сервере (404/410) или это не картинка

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
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                // Ошибка посреди работы: _cache и _source.json не оставляем
                // (при «Стоп» временные файлы убирает внешний код)
                Cleanup(keepCache: false);
                try
                {
                    if (Directory.Exists(outDir) && !Directory.EnumerateFileSystemEntries(outDir).Any())
                        Directory.Delete(outDir);   // пустую папку результата тоже убираем
                }
                catch { }
                throw;
            }
            finally { http.Dispose(); encodeGate.Dispose(); }
        }

        // Общий прогресс: каждая выгрузка занимает свою долю шкалы (внутри: 0–80% скачивание, 80–99% упаковка)
        private void Pct(int local) =>
            percent.Report(Math.Min(99, pctBase + (scopeIndex * 100 + Math.Clamp(local, 0, 100)) / scopeTotal * pctSpan / 100));
        private int pctBase = 0, pctSpan = 100;

        private void Cleanup(bool keepCache)
        {
            if (tempJson != null)
            {
                try { File.Delete(tempJson); } catch { }
            }
            if (!keepCache && Directory.Exists(cacheDir))
            {
                try { Directory.Delete(cacheDir, true); }
                catch (Exception ex) { log.Report("Не удалось удалить временную папку _cache: " + ex.Message); }
            }

            // Если выгружать оказалось нечего, пустую папку результата не оставляем
            try
            {
                if (Directory.Exists(outDir) && !Directory.EnumerateFileSystemEntries(outDir).Any())
                    Directory.Delete(outDir);
            }
            catch { }
        }

        private async Task<string> CoreAsync()
        {
            Directory.CreateDirectory(outDir);
            Directory.CreateDirectory(cacheDir);

            List<Group> groups;
            int offersCount;
            GeoModel geo;

            if (opt.Feeds != null)
            {
                // ---------- Несколько YML-файлов: каждый файл — отдельный магазин ----------
                (groups, offersCount, geo) = await ReadGroupsFeedsAsync();
                pctBase = 40;       // чтение файлов заняло первые 40% полосы прогресса
                pctSpan = 60;
            }
            else
            {
                // ---------- 0. Если указана ссылка — сначала скачиваем файл ----------
                string jsonPath = opt.JsonSource.Trim();
                if (IsUrl(jsonPath))
                {
                    tempJson = Path.Combine(outDir, "_source.json");
                    await DownloadJsonAsync(jsonPath, tempJson, log, ct);
                    jsonPath = tempJson;
                }

                // ---------- 1. Читаем файл и группируем записи по товару ----------
                var format = DetectFormat(jsonPath);
                log.Report(format == SourceFormat.Xml ? "Читаю YML (XML)..." : "Читаю JSON...");
                (groups, offersCount, geo) = format == SourceFormat.Xml
                    ? ReadGroupsXml(jsonPath, ct)
                    : ReadGroups(jsonPath, ct);
            }
            GC.Collect();   // служебные структуры чтения больше не нужны
            ct.ThrowIfCancellationRequested();
            log.Report($"Записей в файле: {offersCount}, уникальных товаров: {groups.Count}");
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
            scopeCount = scopes.Count;
            if (scopes.Count == 0)
                throw new InvalidOperationException("Нет товаров для выгрузки: ни один товар не подошёл под выбранный режим. Причины — в журнале выше.");

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
            if (!missing.IsEmpty)
            {
                log.Report($"Картинок нет на сервере ({missing.Count}), эти товары в Excel не добавлены:");
                foreach (var f in missing) log.Report("  " + f);
            }
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

            if ((mode == RegionMode.AllRegions || mode == RegionMode.PerRegion || mode == RegionMode.RegionsMerged) && geo.UsedRegions.Count == 0)
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
                case RegionMode.RegionsMerged:
                {
                    // Каждый регион проверяем отдельно: процент магазинов считается среди магазинов этого региона,
                    // разброс цен — среди цен этого региона. Товар попадает в общий Excel, если прошёл проверку
                    // в каждом регионе, где он есть. Цену для Excel выбираем по всем его офферам (режим цены из настроек).
                    log.Report($"Режим регионов: один общий Excel, фильтр считается в каждом регионе отдельно — регионов: {geo.UsedRegions.Count}");
                    var failedGroups = new HashSet<Group>();
                    foreach (int r in geo.UsedRegions.OrderBy(x => x))
                    {
                        var cand = groups.Where(g => g.Regions.Contains(r)).ToList();
                        if (cand.Count == 0) continue;

                        var shopSet = new HashSet<int>();
                        if (geo.RegionShops.TryGetValue(r, out var rs))
                            foreach (var sh in rs)
                                if (geo.UsedShops.Contains(sh)) shopSet.Add(sh);

                        bool useShops = opt.ShopFilter && shopSet.Count > 1;
                        int need = useShops ? Math.Max(1, (int)Math.Ceiling(shopSet.Count * opt.ShopPercent / 100m)) : 0;
                        var sub = new Scope { Title = geo.RegionNames[r], Region = r, Candidates = cand, ShopSet = shopSet };
                        var (sel, _) = SelectProducts(sub, useShops, need);

                        var pass = new HashSet<Group>(sel.Select(x => x.Group));
                        foreach (var g in cand)
                            if (!pass.Contains(g)) failedGroups.Add(g);

                        log.Report($"  {geo.RegionNames[r]}: магазинов {shopSet.Count}, товаров {cand.Count}, прошло {pass.Count}" +
                                   (useShops ? $" (нужно минимум в {need} магазинах)" : ""));
                    }

                    var allScope = new Scope();     // без региона и магазина: цены всех офферов
                    var final = new List<(Group Group, decimal Price)>();
                    int noRegion = 0;
                    foreach (var g in groups)
                    {
                        if (g.Regions.Count == 0) { noRegion++; continue; }
                        if (failedGroups.Contains(g)) continue;
                        var prices = ScopePrices(g, allScope);
                        if (prices.Count == 0) continue;
                        final.Add((g, PickPrice(prices, opt.PriceMode)));
                    }

                    log.Report($"Итого в общем Excel: {final.Count} товаров. Отсеяно из-за регионов: {failedGroups.Count}" +
                               (noRegion > 0 ? $", без региона: {noRegion}" : ""));
                    if (final.Count == 0)
                    {
                        log.Report("Ни один товар не прошёл проверку во всех регионах, где он есть. " +
                                   "Уменьшите процент магазинов или увеличьте допустимую разницу цен.");
                        return scopes;
                    }
                    scopes.Add(new Scope
                    {
                        Dir = outDir,
                        Candidates = final.Select(x => x.Group).ToList(),
                        ShopSet = geo.UsedShops,
                        Preselected = final
                    });
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
            if (opt.ShopFilter && scope.Shop < 0 && scope.Preselected == null)   // в выгрузке одного магазина охват не нужен; в «регионах в один Excel» отбор уже сделан
            {
                if (scope.ShopSet.Count > 1)
                {
                    needShops = (int)Math.Ceiling(scope.ShopSet.Count * opt.ShopPercent / 100m);
                    useShops = needShops > 0;
                }
                else L("Магазины (target_shops) не заданы или магазин один — фильтр по магазинам не применён.");
            }

            bool tableShown = false;
            if (useShops && scopeCount <= 3)
            {
                LogCoverageTable(scope, needShops, L);
                tableShown = true;
            }

            var (selected, st) = SelectProducts(scope, useShops, needShops);

            if (useShops && selected.Count == 0)
            {
                if (!tableShown) LogCoverageTable(scope, needShops, L);
                // Ничего не подошло: не подставляем «все товары» и ничего не скачиваем, а останавливаемся с понятной ошибкой
                int best = scope.Candidates.Count == 0 ? 0 : scope.Candidates.Max(g => CountIn(g.Shops, scope.ShopSet));
                string where = scope.Title.Length > 0 ? $" ({scope.Kind}: {scope.Title})" : "";
                if (st.ShopDropped == scope.Candidates.Count)
                    throw new InvalidOperationException(
                        $"Ни один товар не встречается минимум в {opt.ShopPercent}% магазинов ({needShops} из {scope.ShopSet.Count}){where}. " +
                        $"Максимальный охват — {best} из {scope.ShopSet.Count} магазинов. Уменьшите процент магазинов. " +
                        "Ничего не скачано.");
                throw new InvalidOperationException(
                    $"Товары, которые есть минимум в {opt.ShopPercent}% магазинов, нашлись ({scope.Candidates.Count - st.ShopDropped}), " +
                    $"но у всех разброс цен между магазинами больше {opt.MaxSpreadPercent}%{where}. " +
                    "Увеличьте допустимый разброс цен или уменьшите процент магазинов. Ничего не скачано.");
            }

            if (useShops)
                L($"Фильтр по магазинам: товар должен быть минимум в {needShops} из {scope.ShopSet.Count} ({opt.ShopPercent}%); " +
                  $"отсеяно: {st.ShopDropped}");
            if (scope.Preselected != null)
            {
                L($"Взято товаров: {selected.Count}");
            }
            else
            {
                L($"Порог разброса цен: {opt.MaxSpreadPercent}%");
                L($"Взято товаров: {selected.Count} " +
                  $"(цена одинаковая: {st.Equal}, разница в пределах порога: {st.Similar}, единственная цена: {st.Single})");
                L($"Пропущено из-за разброса цен: {st.Dropped}");
                if (st.NoPrice > 0) L($"Пропущено без корректной цены: {st.NoPrice}");
            }

            if (selected.Count == 0)
                throw new InvalidOperationException(
                    "Нет товаров для выгрузки" + (scope.Title.Length > 0 ? $" ({scope.Kind}: {scope.Title})" : "") + ". " +
                    (st.Dropped > 0
                        ? $"У всех товаров ({st.Dropped}) разброс цен больше {opt.MaxSpreadPercent}%. Увеличьте допустимую разницу цен."
                        : "Подходящих товаров с ценой не нашлось."));

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

                cur.Rows.Add(new Row(g.Description, file, price, g.DateStart, g.DateEnd, g.OldPrice));
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
            // Колонка price_old добавляется в конец, только если у товаров есть старая цена (YML);
            // для JSON формат таблицы остаётся прежним
            bool hasOld = part.Rows.Any(r => r.OldPrice.HasValue);
            var headers = new List<string> { "description", "image", "price_new", "date_start", "date_end" };
            if (hasOld) headers.Add("price_old");
            for (int c = 0; c < headers.Count; c++) ws.Cell(1, c + 1).Value = headers[c];
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
                if (hasOld && r.OldPrice.HasValue) ws.Cell(row, 6).Value = r.OldPrice.Value;
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
            if (hasOld) ws.Column(6).Width = 12;
            wb.SaveAs(path);
        }

        // Таблица для подбора настроек: сколько товаров проходит при разном проценте магазинов
        // и сколько из них с одной ценой (цена совпадает во всех магазинах) / с разбросом в допустимых пределах.
        private void LogCoverageTable(Scope scope, int needShops, Action<string> L)
        {
            int total = scope.ShopSet.Count;
            if (total <= 1 || scope.Candidates.Count == 0) return;

            decimal limit = opt.MaxSpreadPercent / 100m;
            var rows = new List<(int Cov, bool Equal, bool InSpread)>(scope.Candidates.Count);
            foreach (var g in scope.Candidates)
            {
                var prices = ScopePrices(g, scope);
                if (prices.Count == 0) continue;
                decimal min = prices.Min(), max = prices.Max();
                rows.Add((CountIn(g.Shops, scope.ShopSet), max == min, (max - min) / min <= limit));
            }

            var percents = new SortedSet<decimal>(new decimal[] { 100, 95, 90, 80, 70, 50, 30, 10, opt.ShopPercent });
            L($"Охват по магазинам (всего магазинов: {total}, товаров с ценой: {rows.Count}). " +
              $"Допустимая разница цен: {opt.MaxSpreadPercent}%:");
            foreach (var p in percents.Reverse())
            {
                int need = Math.Max(1, (int)Math.Ceiling(total * p / 100m));
                int pass = 0, eq = 0, sp = 0;
                foreach (var r in rows)
                {
                    if (r.Cov < need) continue;
                    pass++;
                    if (r.Equal) eq++;
                    if (r.InSpread) sp++;
                }
                string mark = p == opt.ShopPercent ? "  ◄ выбрано" : "";
                L($"   минимум {p,3:0.##}% ({need} из {total}): товаров {pass}, с одной ценой {eq}, " +
                  $"в пределах разброса {sp}{mark}");
            }
        }

        // ---------- Отбор товаров внутри одной выгрузки ----------
        private (List<(Group Group, decimal Price)> Items, SelStats Stats) SelectProducts(
            Scope scope, bool useShops, int needShops)
        {
            if (scope.Preselected != null) return (scope.Preselected, new SelStats());

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
        // Регион = субъект РФ: включено в настройках и справочник найден
        private bool BySubject => opt.RegionBySubject && RuGeo.Loaded;
        private int scopeCount;         // сколько выгрузок (Excel) получится: таблицу охвата печатаем, только если их мало

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
                                missing.Add($"{pair.Key} -> {ex.Message}");   // не картинка — повторять бессмысленно
                                break;
                            }
                            catch (HttpRequestException ex) when (
                                ex.StatusCode is System.Net.HttpStatusCode.NotFound or System.Net.HttpStatusCode.Gone)
                            {
                                missing.Add($"{pair.Key} -> файла нет на сервере ({(int)ex.StatusCode!.Value})");
                                break;                                        // 404/410 — повторы и докачка не помогут
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

        // ---------- Чтение YML (XML) ----------
        // Формат Яндекс.Маркета / Едадила: <offer id="..."><name/><price/><oldprice/><picture/>...</offer>.
        // Читается потоком, по одному офферу.
        private readonly record struct YmlOffer(string Id, string Name, string? Image, decimal Price, decimal? OldPrice);

        private sealed class YmlStats
        {
            public int Count, NotPromo, Skipped;
        }

        // Читает один YML-файл и вызывает onOffer для каждого подходящего оффера
        private YmlStats ReadYmlFile(string path, Action<YmlOffer> onOffer, CancellationToken ct, Action<int>? onPercent)
        {
            var st = new YmlStats();

            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16,
                                          FileOptions.SequentialScan);
            long length = Math.Max(1, fs.Length);
            using var reader = XmlReader.Create(fs, CreateXmlSettings());

            string? shopUrl = null;
            int lastPercent = -1;

            reader.MoveToContent();
            while (!reader.EOF)
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    // Адрес магазина нужен, чтобы достроить относительные ссылки на картинки
                    if (reader.Depth == 2 && reader.Name == "url")
                    {
                        shopUrl = reader.ReadElementContentAsString().Trim();
                        continue;
                    }

                    if (reader.Name == "offer")
                    {
                        var o = (XElement)XNode.ReadFrom(reader);   // после чтения reader уже стоит на следующем узле
                        st.Count++;

                        string name = (o.Element("name")?.Value ?? "").Trim();
                        string id = (o.Attribute("id")?.Value ?? "").Trim();

                        var price = ParseYmlPrice(o.Element("price")?.Value);
                        var oldPrice = ParseYmlPrice(o.Element("oldprice")?.Value);
                        bool promo = oldPrice.HasValue && oldPrice.Value > 0;

                        if (id.Length == 0 && name.Length == 0) st.Skipped++;
                        else if (opt.PromoOnly && !promo) st.NotPromo++;       // режим «только акции»: нужен oldprice
                        else
                            onOffer(new YmlOffer(id, name, FirstPicture(o, shopUrl),
                                                 price.HasValue && price.Value > 0 ? price.Value : 0m,
                                                 promo ? oldPrice : null));

                        if ((st.Count & 255) == 0)
                        {
                            ct.ThrowIfCancellationRequested();
                            if (onPercent != null)
                            {
                                int p = (int)(fs.Position * 100 / length);
                                if (p != lastPercent) { lastPercent = p; onPercent(p); }
                            }
                        }
                        continue;
                    }
                }
                reader.Read();
            }
            return st;
        }

        // Добавляет оффер в группу товара (создаёт группу при необходимости); cat — код каталога/файла (-1 — нет)
        private static void AddYmlToGroups(List<Group> list, Dictionary<string, Group> index, string key, YmlOffer o, int cat)
        {
            if (!index.TryGetValue(key, out var grp))
            {
                grp = new Group();
                index[key] = grp;
                list.Add(grp);
            }

            if (grp.Description.Length == 0) grp.Description = o.Name;
            if (string.IsNullOrWhiteSpace(grp.ImageUrl)) grp.ImageUrl = o.Image;
            if (o.OldPrice.HasValue) grp.OldPrice ??= o.OldPrice;
            grp.Offers.Add(new PriceEntry(o.Price, cat));
        }

        // Один YML-файл: магазинов и регионов в нём нет — гео задаётся в программе.
        private (List<Group> Groups, int OffersCount, GeoModel Geo) ReadGroupsXml(string path, CancellationToken ct)
        {
            var list = new List<Group>();
            var index = new Dictionary<string, Group>();

            var geo = new GeoModel();
            currentGeo = geo;
            if (opt.RegionMode != RegionMode.Ignore || opt.ShopFilter)
                log.Report("В YML нет данных о магазинах и регионах — разбивка по гео и фильтр по магазинам не применяются.");

            int nextLog = 10;
            var st = ReadYmlFile(path,
                o => AddYmlToGroups(list, index, o.Id.Length > 0 ? o.Id : o.Name, o, -1),
                ct,
                p =>
                {
                    if (p >= nextLog)
                    {
                        log.Report($"Чтение товаров: {p}%");
                        nextLog = p / 10 * 10 + 10;
                    }
                });

            if (st.Count == 0) throw new Exception("В файле нет ни одного <offer>. Это точно YML-фид?");
            if (opt.PromoOnly)
                log.Report($"Только акционные позиции: пропущено без oldprice — {st.NotPromo} из {st.Count}");
            if (st.Skipped > 0)
                log.Report($"Пропущено офферов без id и названия: {st.Skipped}");

            foreach (var g in list) g.Offers.TrimExcess();
            return (list, st.Count, geo);
        }

        // Несколько файлов в один результат (YML и/или JSON).
        // YML: каждый файл — отдельный магазин (как каталог в JSON).
        // JSON: магазины берутся из каталогов (target_shops) внутри файла; если файл вообще без гео — он один магазин.
        // Один и тот же магазин в разных файлах учитывается один раз, один и тот же товар — тоже один раз.
        // Дальше работают обычные механизмы: охват магазинов, разброс цен, один общий Excel и один набор картинок.
        private async Task<(List<Group> Groups, int OffersCount, GeoModel Geo)> ReadGroupsFeedsAsync()
        {
            var feeds = opt.Feeds!;
            WarnNoGeoData();
            var s = new ReadState();
            var list = s.List;
            var index = s.Index;
            var geo = s.Geo;
            currentGeo = geo;
            var shopIndex = s.ShopIndex;
            var regionIndex = s.RegionIndex;
            var catShops = s.CatShops;
            var catRegions = s.CatRegions;
            var usedLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var failedFeeds = new List<string>();

            string tmpFeed = Path.Combine(outDir, "_feed.tmp");
            tempJson = tmpFeed;                      // Cleanup удалит временный файл при любом исходе

            int total = feeds.Count, okFeeds = 0, totalOffers = 0, totalNotPromo = 0, totalDupShops = 0, totalSkipped = 0;
            log.Report($"Режим объединения: файлов в списке: {total}");

            PreCheckJson(feeds);

            for (int i = 0; i < total; i++)
            {
                ct.ThrowIfCancellationRequested();
                var feed = feeds[i];
                percent.Report(Math.Min(39, i * 40 / total));

                // Название должно быть уникальным, иначе два файла склеятся в один магазин
                string label = string.IsNullOrWhiteSpace(feed.Label) ? $"feed_{i + 1}" : feed.Label.Trim();
                string unique = label;
                int k = 2;
                while (!usedLabels.Add(unique)) unique = $"{label} ({k++})";

                // Товары файла собираем отдельно и добавляем в общий набор только если файл прочитан целиком
                var buffer = new List<(string Key, YmlOffer Offer)>();
                YmlStats? st = null;
                JsonFileStats? js = null;
                try
                {
                    string path = feed.Source.Trim();
                    if (IsUrl(path))
                    {
                        for (int attempt = 1; ; attempt++)
                        {
                            try
                            {
                                await DownloadJsonAsync(path, tmpFeed, log, ct, quiet: true);
                                break;
                            }
                            catch (Exception) when (attempt < 3 && !ct.IsCancellationRequested)
                            {
                                await Task.Delay(1000 * attempt, ct);
                            }
                        }
                        path = tmpFeed;
                    }
                    else if (!File.Exists(path))
                    {
                        throw new FileNotFoundException("файл не найден");
                    }

                    if (DetectFormat(path) == SourceFormat.Json)
                    {
                        js = ReadJsonFile(s, path, unique, feed, unique, atomic: true, null);
                    }
                    else
                    {
                        st = ReadYmlFile(path, o =>
                        {
                            // Один и тот же товар в разных магазинах: совпадают id и название
                            string key = o.Id.Length > 0 ? o.Id + "\u0001" + NormalizeText(o.Name) : NormalizeText(o.Name);
                            buffer.Add((key, o));
                        }, ct, null);

                        if (st.Count == 0) throw new InvalidDataException("в файле нет офферов");
                    }
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    log.Report($"[{i + 1}/{total}] {unique}: ПРОПУЩЕН — {ex.Message}");
                    failedFeeds.Add(feed.Region == null ? $"{feed.Source};{feed.Label}" : $"{feed.Source};{feed.Label};{feed.Region}");
                    usedLabels.Remove(unique);
                    continue;
                }
                finally
                {
                    try { if (File.Exists(tmpFeed)) File.Delete(tmpFeed); } catch { }
                }

                // ---------- JSON: файл уже добавлен в общий набор ----------
                if (js != null)
                {
                    okFeeds++;
                    totalOffers += js.Offers;
                    totalNotPromo += js.NotPromo;
                    totalDupShops += js.DuplicateCatalogs;
                    totalSkipped += js.SkippedDuplicate;

                    string what = js.WholeFileIsShop
                        ? "в файле нет данных о магазинах — весь файл считается одним магазином"
                        : js.Catalogs == 0 && js.DuplicateCatalogs > 0
                            ? "файл целиком повторяет уже прочитанные"
                            : $"магазинов {js.Catalogs}";
                    if (js.DuplicateCatalogs > 0)
                        what += $", повторных (уже были в других файлах) {js.DuplicateCatalogs} — пропущены";
                    log.Report($"[{i + 1}/{total}] {unique}: {what}, офферов {js.Offers}, новых товаров {js.NewProducts}");

                    if (opt.Economy) GC.Collect();   // слабый ПК: буфер файла больше не нужен
                    continue;
                }

                // ---------- YML: файл — это новый «каталог» с одним магазином (и, если задан, регионом) ----------
                int cat = catShops.Count;
                int shop = Intern(shopIndex, geo.ShopNames, unique);
                var regions = new HashSet<int>();
                if (!string.IsNullOrWhiteSpace(feed.Region))
                    regions.Add(Intern(regionIndex, geo.RegionNames, feed.Region.Trim()));

                catShops.Add(new[] { shop });
                catRegions.Add(regions.ToArray());
                foreach (var r in regions)
                {
                    if (!geo.RegionShops.TryGetValue(r, out var rs)) geo.RegionShops[r] = rs = new HashSet<int>();
                    rs.Add(shop);
                }

                foreach (var (key, offer) in buffer) AddYmlToGroups(list, index, key, offer, cat);

                okFeeds++;
                totalOffers += st!.Count;
                totalNotPromo += st.NotPromo;
                log.Report($"[{i + 1}/{total}] {unique}: товаров {st.Count}");

                if (opt.Economy && (i & 31) == 31) GC.Collect();   // слабый ПК: не даём куче разрастаться
            }

            if (okFeeds == 0) throw new Exception("Ни один файл списка не удалось прочитать.");

            if (failedFeeds.Count > 0)
            {
                string failedPath = Path.Combine(outDir, "failed_feeds.txt");
                File.WriteAllLines(failedPath, failedFeeds);
                log.Report($"⚠ Не прочитано файлов: {failedFeeds.Count} из {total}. Они не участвуют в охвате магазинов. " +
                           $"Список для повторной загрузки: {failedPath}");
            }
            if (opt.PromoOnly)
                log.Report($"Только акционные позиции: пропущено без oldprice / price_old — {totalNotPromo} из {totalOffers}");
            if (totalDupShops > 0)
                log.Report($"Повторяющихся магазинов пропущено: {totalDupShops} (офферов в них: {totalSkipped}). " +
                           "Каждый магазин и каждый товар учтены один раз.");

            // Магазины и регионы товара: по файлам и каталогам, в которых он встретился
            FinishGeo(s);
            ReportDerivedRegions(s);
            log.Report($"Прочитано файлов: {okFeeds}, магазинов: {geo.UsedShops.Count}" +
                       (geo.UsedRegions.Count > 0 ? $", регионов: {geo.UsedRegions.Count}" : "") +
                       $", уникальных товаров: {list.Count}");

            return (list, totalOffers, geo);
        }

        // Первая непустая картинка оффера (в YML их может быть несколько)
        private static string? FirstPicture(XElement offer, string? shopUrl)
        {
            foreach (var pic in offer.Elements("picture"))
            {
                var url = ResolvePictureUrl(pic.Value, shopUrl);
                if (url != null) return url;
            }
            return null;
        }

        // ---------- Чтение JSON ----------
        // JSON читается потоком (Utf8JsonReader): в памяти одновременно только один оффер, поэтому расход
        // памяти зависит от числа уникальных товаров, а не от размера файла. Два прохода по файлу:
        // сначала каталоги (география), затем офферы.
        // Один и тот же код читает и одиночный файл, и каждый файл в режиме объединения.

        // Всё, что накапливается при чтении: товары, магазины, регионы, каталоги.
        private sealed class ReadState
        {
            public readonly List<Group> List = new();
            public readonly Dictionary<string, Group> Index = new();
            public readonly GeoModel Geo = new();
            public readonly Dictionary<string, int> ShopIndex = new(StringComparer.OrdinalIgnoreCase);
            public readonly Dictionary<string, int> RegionIndex = new(StringComparer.OrdinalIgnoreCase);
            public readonly List<int[]> CatRegions = new();
            public readonly List<int[]> CatShops = new();
            // Каталоги, уже прочитанные из предыдущих файлов: ключ каталога -> имя файла
            public readonly Dictionary<string, string> CatalogOwners = new(StringComparer.OrdinalIgnoreCase);
            public int DerivedCatalogs;                       // у скольких каталогов регион определён по адресу магазина
            public readonly HashSet<string> UnparsedAddresses = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class JsonFileStats
        {
            public int Catalogs, DuplicateCatalogs, Offers, NotPromo, SkippedDuplicate, NewProducts;
            public bool WholeFileIsShop;     // в файле нет ни магазинов, ни регионов: весь файл считаем одним магазином
        }

        private readonly record struct BufferedOffer(
            string Key, string? Description, string? Image, DateTime? DateStart, DateTime? DateEnd, decimal Price, int Cat);

        // Магазины (target_shops) и регионы (target_regions, иначе region) одного каталога.
        // Если регионов нет, а режим регионов включён (deriveFromAddress), регионом считается город из адреса магазина.
        // derived = true, если регионы получены именно так; unparsed получает адреса, из которых город определить не удалось.
        private static void ReadCatalogGeo(JsonElement c, bool deriveFromAddress, bool bySubject, List<string>? unparsed,
                                           out List<string> shops, out List<string> regions, out bool derived)
        {
            derived = false;
            regions = GetStringList(c, "target_regions").Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (regions.Count == 0)
            {
                string? one = GetString(c, "region");
                if (!string.IsNullOrWhiteSpace(one)) regions.Add(one.Trim());
            }
            shops = GetStringList(c, "target_shops").Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            // Названия из target_regions приводим к официальным по справочнику («Московская обл.» -> «Московская область»);
            // в режиме субъектов РФ город из target_regions заменяется его регионом
            if (regions.Count > 0 && RuGeo.Loaded)
                regions = regions.Select(x => RuGeo.NormalizeRegion(x, bySubject) ?? x)
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            if (regions.Count == 0 && deriveFromAddress && shops.Count > 0)
            {
                foreach (var sh in shops)
                {
                    string? city = bySubject ? RuGeo.RegionOfAddress(sh) : CityFromAddress(sh);
                    if (city == null) unparsed?.Add(sh);
                    else if (!regions.Contains(city, StringComparer.OrdinalIgnoreCase)) regions.Add(city);
                }
                derived = true;
            }
        }

        private static readonly Regex SettlementPrefixRx = new(
            @"^(?:г|город|пгт|рп|р\.п|пос|посёлок|поселок|п|с|село|ст-ца|станица|х|хутор|аул|дер|деревня)(?:\.\s*|\s+)(?<n>\p{L}.*)$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SettlementSuffixRx = new(
            @"^(?<n>\p{L}.*?)\s+(?:г|гор)\.?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex SkipPartRx = new(
            @"^(?:россия|рф|russia|\d{5,6}|\d.*)$|обл(?:\.|асть|\b)|край\b|респ(?:\.|ублика|\b)|автоном|\bао\b|округ|район|р-н",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex StreetPartRx = new(
            @"^(?:ул|улица|пр|пр-т|просп|проспект|пер|переулок|б-р|бульвар|ш|шоссе|наб|набережная|пл|площадь|тракт|проезд|аллея|линия|мкр|микрорайон|тц|трц|тк|д|дом|к|корп|стр|вл|кв|пом|литер|лит)(?:\.|\s|$)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>Город из адреса: «г. Москва, пр. Ленинградский, д. 77» -> «Москва». null, если определить не удалось.</summary>
        public static string? CityFromAddress(string address)
        {
            var parts = address.Split(',').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
            if (parts.Count == 0) return null;

            // 1. Явный населённый пункт: «г. Химки», «город Химки», «г Химки», «Химки г.», «пгт Янтарный»
            foreach (var p in parts)
            {
                var m = SettlementPrefixRx.Match(p);
                if (!m.Success) m = SettlementSuffixRx.Match(p);
                if (m.Success) return CleanCity(m.Groups["n"].Value);
            }

            // 2. Без пометки («Москва, ул. Ленина, 5»): первая часть, которая не страна, не индекс, не область и не улица
            foreach (var p in parts)
            {
                if (SkipPartRx.IsMatch(p) || StreetPartRx.IsMatch(p)) continue;
                return CleanCity(p);
            }
            return null;
        }

        private static string? CleanCity(string s)
        {
            s = Regex.Replace(s, @"\s+", " ").Trim(' ', '.', ',', ';');
            return s.Length == 0 ? null : s;
        }

        // Каталог из другого файла считается повтором, только если совпадает всё: id, магазины и регионы.
        // Без id, магазинов и регионов каталог повтором не считается никогда.
        private static string? BuildCatalogKey(string? id, List<string> shops, List<string> regions)
        {
            if (string.IsNullOrWhiteSpace(id) && shops.Count == 0 && regions.Count == 0) return null;
            return (id?.Trim() ?? "") + "\u0001" +
                   string.Join("|", shops.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)) + "\u0001" +
                   string.Join("|", regions.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
        }

        // Читает один JSON-файл в общее состояние.
        // atomic = true (режим объединения): товары файла сначала копятся отдельно и попадают в общий набор только
        // после чтения всего файла, поэтому битый файл ничего не портит и просто пропускается.
        // atomic = false (один файл): товары сразу идут в общий набор — так меньше памяти на очень больших файлах.
        private JsonFileStats ReadJsonFile(ReadState s, string path, string fileName, FeedSource? feed,
                                           string? label, bool atomic, Action<int>? onPercent)
        {
            var st = new JsonFileStats();
            var geo = s.Geo;

            int baseCat = s.CatShops.Count;       // индекс первого каталога этого файла в общем списке
            int baseMulti = geo.Multi.Count;      // индекс первого «составного» кода (оффер в нескольких каталогах)

            var localCats = new List<(string[] Shops, string[] Regions, bool Derived)>();
            bool deriveRegions = opt.RegionMode is RegionMode.AllRegions or RegionMode.PerRegion or RegionMode.RegionsMerged;
            var unparsed = new List<string>();
            var localMulti = new List<int[]>();           // коды каталогов уже в общей нумерации
            var ownedKeys = new List<string>();
            var offerCats = new Dictionary<ulong, int>(); // хэш id оффера -> каталог (общая нумерация) или составной код
            HashSet<ulong>? skipped = null;               // офферы каталогов, которые уже были в предыдущих файлах
            bool anyGeo = false;

            // Проход 1: каталоги
            ScanArray(path, "catalogs", c =>
            {
                ReadCatalogGeo(c, deriveRegions, BySubject, unparsed, out var shops, out var regions, out var derived);
                string? key = BuildCatalogKey(GetString(c, "id"), shops, regions);
                bool dup = key != null && s.CatalogOwners.ContainsKey(key);

                int ci = baseCat + localCats.Count;
                if (dup)
                {
                    st.DuplicateCatalogs++;
                }
                else
                {
                    localCats.Add((shops.ToArray(), regions.ToArray(), derived));
                    if (derived) s.DerivedCatalogs++;
                    if (key != null) ownedKeys.Add(key);
                    if (shops.Count > 0 || regions.Count > 0) anyGeo = true;
                }

                if (c.TryGetProperty("offers", out var ids) && ids.ValueKind == JsonValueKind.Array)
                {
                    foreach (var idEl in ids.EnumerateArray())
                    {
                        if (idEl.ValueKind != JsonValueKind.String) continue;
                        string? oid = idEl.GetString();
                        if (string.IsNullOrEmpty(oid)) continue;

                        ulong h = HashId(oid);

                        if (dup)
                        {
                            (skipped ??= new HashSet<ulong>()).Add(h);
                            continue;
                        }

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
                            localMulti.Add(new[] { code, ci });
                            offerCats[h] = -(baseMulti + localMulti.Count - 1) - 2;
                        }
                        else
                        {
                            int mi = -code - 2 - baseMulti;
                            var old = localMulti[mi];
                            if (Array.IndexOf(old, ci) < 0)
                            {
                                var arr = new int[old.Length + 1];
                                old.CopyTo(arr, 0);
                                arr[^1] = ci;
                                localMulti[mi] = arr;
                            }
                        }
                    }
                }
            }, ct);

            // В режиме объединения файл без всякой географии — это один магазин (как YML-файл)
            int? wholeFileCat = null;
            if (feed != null && !anyGeo && st.DuplicateCatalogs == 0)
            {
                st.WholeFileIsShop = true;
                st.DuplicateCatalogs = 0;
                skipped = null;
                localCats.Clear();
                localMulti.Clear();
                ownedKeys.Clear();
                localCats.Add((new[] { (label ?? fileName).Trim() },
                               string.IsNullOrWhiteSpace(feed.Region) ? Array.Empty<string>() : new[] { feed.Region.Trim() }, false));
                wholeFileCat = baseCat;
            }
            st.Catalogs = localCats.Count;
            foreach (var u in unparsed) s.UnparsedAddresses.Add(u);

            // Каталоги переносим в общее состояние (магазины и регионы получают общие номера)
            void CommitCatalogs()
            {
                foreach (var m in localMulti) geo.Multi.Add(m);
                foreach (var (shopNames, regionNames, derived) in localCats)
                {
                    var shops = shopNames.Select(n => Intern(s.ShopIndex, geo.ShopNames, n)).Distinct().ToArray();
                    var regions = regionNames.Select(n => Intern(s.RegionIndex, geo.RegionNames, n)).Distinct().ToArray();
                    s.CatShops.Add(shops);
                    s.CatRegions.Add(regions);

                    if (derived)
                    {
                        // Регион получен из адресов: каждый магазин относим только к городу из его собственного адреса
                        foreach (var sn in shopNames)
                        {
                            string? city = BySubject ? RuGeo.RegionOfAddress(sn) : CityFromAddress(sn);
                            if (city == null) continue;
                            int r = Intern(s.RegionIndex, geo.RegionNames, city);
                            if (!geo.RegionShops.TryGetValue(r, out var rs)) geo.RegionShops[r] = rs = new HashSet<int>();
                            rs.Add(Intern(s.ShopIndex, geo.ShopNames, sn));
                        }
                        continue;
                    }
                    foreach (var r in regions)
                    {
                        if (!geo.RegionShops.TryGetValue(r, out var rs)) geo.RegionShops[r] = rs = new HashSet<int>();
                        rs.UnionWith(shops);
                    }
                }
                foreach (var k in ownedKeys) s.CatalogOwners.TryAdd(k, fileName);
            }

            if (!atomic) CommitCatalogs();

            // Проход 2: офферы
            var buffer = atomic ? new List<BufferedOffer>() : null;
            var localSeen = atomic ? new HashSet<string>() : null;   // у первого оффера товара в файле берём текст и картинку

            bool found = ScanArray(path, "offers", o =>
            {
                st.Offers++;

                // Режим «только акции»: оффер без корректной старой цены (price_old) не учитываем совсем —
                // ни в ценах, ни в охвате магазинов и регионов
                if (opt.PromoOnly)
                {
                    var oldPrice = GetDecimal(o, "price_old");
                    if (!(oldPrice.HasValue && oldPrice.Value > 0)) { st.NotPromo++; return; }
                }

                string id = GetString(o, "id") ?? "";
                ulong h = id.Length > 0 ? HashId(id) : 0UL;

                // Оффер каталога, который уже был в предыдущем файле, второй раз не берём
                if (skipped != null && id.Length > 0 && skipped.Contains(h) && !offerCats.ContainsKey(h))
                {
                    st.SkippedDuplicate++;
                    return;
                }

                // id вида "<код товара>_<регион/магазин>" -> ключ товара = часть до первого "_".
                // Если "_" нет (как в palich), каждая запись — отдельный товар.
                string key;
                int us = id.IndexOf('_');
                if (us > 0) key = id[..us];
                else key = id.Length > 0 ? id : (GetString(o, "description") ?? "");

                int cat = wholeFileCat ?? ((id.Length > 0 && offerCats.TryGetValue(h, out var cc)) ? cc : -1);
                var price = GetDecimal(o, "price_new");
                decimal p = price.HasValue && price.Value > 0 ? price.Value : 0m;

                if (buffer == null)
                {
                    if (!s.Index.TryGetValue(key, out var grp))
                    {
                        grp = new Group();
                        s.Index[key] = grp;
                        s.List.Add(grp);
                        st.NewProducts++;
                    }

                    // Текст, ссылку и даты читаем только пока они не найдены — лишних строк не создаём
                    if (grp.Description.Length == 0) grp.Description = GetString(o, "description") ?? "";
                    if (string.IsNullOrWhiteSpace(grp.ImageUrl)) grp.ImageUrl = GetString(o, "image");
                    if (!grp.DateStart.HasValue) grp.DateStart = GetDate(o, "date_start");
                    if (!grp.DateEnd.HasValue) grp.DateEnd = GetDate(o, "date_end");
                    grp.Offers.Add(new PriceEntry(p, cat));
                }
                else
                {
                    string? desc = null, img = null;
                    DateTime? ds = null, de = null;
                    if (localSeen!.Add(key))
                    {
                        desc = GetString(o, "description");
                        img = GetString(o, "image");
                        ds = GetDate(o, "date_start");
                        de = GetDate(o, "date_end");
                    }
                    buffer.Add(new BufferedOffer(key, desc, img, ds, de, p, cat));
                }
            }, ct, onPercent);

            if (!found) throw new Exception("В JSON нет массива \"offers\".");

            if (atomic)
            {
                if (st.Offers == 0) throw new InvalidDataException("в файле нет офферов");

                CommitCatalogs();
                foreach (var b in buffer!)
                {
                    if (!s.Index.TryGetValue(b.Key, out var grp))
                    {
                        grp = new Group();
                        s.Index[b.Key] = grp;
                        s.List.Add(grp);
                        st.NewProducts++;
                    }
                    if (grp.Description.Length == 0 && b.Description != null) grp.Description = b.Description;
                    if (string.IsNullOrWhiteSpace(grp.ImageUrl)) grp.ImageUrl = b.Image;
                    if (!grp.DateStart.HasValue) grp.DateStart = b.DateStart;
                    if (!grp.DateEnd.HasValue) grp.DateEnd = b.DateEnd;
                    grp.Offers.Add(new PriceEntry(b.Price, b.Cat));
                }
                buffer!.Clear();
            }

            return st;
        }

        // Магазины и регионы товара считаем один раз по уникальным каталогам его офферов.
        // Учитываем только магазины и регионы, в которых реально есть товары (пустые каталоги не считаем).
        private void FinishGeo(ReadState s)
        {
            s.Geo.CatShops = s.CatShops.ToArray();
            s.Geo.CatRegions = s.CatRegions.ToArray();
            currentGeo = s.Geo;

            foreach (var g in s.List)
            {
                ct.ThrowIfCancellationRequested();
                var seen = new HashSet<int>();
                foreach (var e in g.Offers)
                    if (e.Cat != -1 && seen.Add(e.Cat)) AddGeo(g, s.Geo, e.Cat);

                g.Offers.TrimExcess();
                s.Geo.UsedShops.UnionWith(g.Shops);
                s.Geo.UsedRegions.UnionWith(g.Regions);
            }
        }

        // Сообщает, если регионы определялись по адресам магазинов
        private void ReportDerivedRegions(ReadState s)
        {
            if (s.DerivedCatalogs == 0) return;
            log.Report($"Регионы определены по адресам магазинов (target_regions в файлах нет): найдено {s.Geo.RegionNames.Count} " +
                       (BySubject ? "(субъекты РФ по справочнику)." : "(города)."));
            if (s.UnparsedAddresses.Count > 0)
                log.Report($"⚠ Не удалось определить {(BySubject ? "регион" : "город")} у {s.UnparsedAddresses.Count} адресов, они не попадут ни в один регион. " +
                           $"Например: {string.Join(" | ", s.UnparsedAddresses.Take(3))}");
        }

        private void WarnNoGeoData()
        {
            if (opt.RegionBySubject && !RuGeo.Loaded)
                log.Report("⚠ Справочник регионов (ru_geo.tsv) не встроен в программу: регионами останутся города из адресов. " +
                           "Добавьте ru_geo.tsv в проект как EmbeddedResource.");
        }

        private (List<Group> Groups, int OffersCount, GeoModel Geo) ReadGroups(string jsonPath, CancellationToken ct)
        {
            WarnNoGeoData();
            var s = new ReadState();
            currentGeo = s.Geo;

            int nextLog = 10;
            var st = ReadJsonFile(s, jsonPath, Path.GetFileName(jsonPath), null, null, atomic: false, p =>
            {
                if (p >= nextLog)
                {
                    log.Report($"Чтение товаров: {p}%");
                    nextLog = p / 10 * 10 + 10;
                }
            });

            if (opt.PromoOnly)
                log.Report($"Только акционные позиции: пропущено без price_old — {st.NotPromo} из {st.Offers}");

            FinishGeo(s);
            ReportDerivedRegions(s);
            return (s.List, st.Offers, s.Geo);
        }

        // Быстрая проверка локальных JSON перед долгим чтением: читаются только каталоги (они обычно в начале файла),
        // поэтому это занимает секунды. Сообщает о битых файлах, о повторяющихся магазинах и о числе магазинов.
        // Ссылки не проверяются: ради этого пришлось бы скачивать их дважды.
        private void PreCheckJson(IReadOnlyList<FeedSource> feeds)
        {
            var owners = new Dictionary<string, (int Index, string Name)>(StringComparer.OrdinalIgnoreCase);
            var shopNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var regionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int files = 0, bad = 0, cats = 0, dups = 0;
            bool deriveRegions = opt.RegionMode is RegionMode.AllRegions or RegionMode.PerRegion or RegionMode.RegionsMerged;

            for (int i = 0; i < feeds.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                string path = feeds[i].Source.Trim();
                if (IsUrl(path) || !File.Exists(path)) continue;

                try { if (DetectFormat(path) != SourceFormat.Json) continue; }
                catch { continue; }

                files++;
                int fileNo = i;
                string name = Path.GetFileName(path);
                int fileCats = 0, fileDups = 0;
                var firstFiles = new HashSet<string>();

                try
                {
                    bool found = ScanArray(path, "catalogs", c =>
                    {
                        ReadCatalogGeo(c, deriveRegions, BySubject, null, out var shops, out var regions, out _);
                        string? key = BuildCatalogKey(GetString(c, "id"), shops, regions);
                        fileCats++;
                        if (key != null)
                        {
                            if (owners.TryGetValue(key, out var first))
                            {
                                if (first.Index != fileNo)     // повтор внутри одного файла повтором не считаем
                                {
                                    fileDups++;
                                    firstFiles.Add(first.Name);
                                }
                            }
                            else owners[key] = (fileNo, name);
                        }
                        foreach (var sh in shops) shopNames.Add(sh);
                        foreach (var rg in regions) regionNames.Add(rg);
                    }, ct);

                    if (!found)
                        log.Report($"  Проверка: {name} — нет массива catalogs, файл будет считаться одним магазином.");
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    bad++;
                    log.Report($"  ⚠ Проверка: {name} — не читается ({ex.Message}). Файл будет пропущен.");
                    continue;
                }

                cats += fileCats;
                dups += fileDups;
                if (fileDups > 0)
                    log.Report($"  Проверка: {name} — повторяющихся магазинов {fileDups} из {fileCats} " +
                               $"(уже есть в: {string.Join(", ", firstFiles.Take(3))}); они будут пропущены.");
            }

            if (files == 0) return;

            log.Report($"Проверка JSON: файлов {files}, каталогов {cats}, разных магазинов {shopNames.Count}" +
                       (regionNames.Count > 0 ? $", регионов {regionNames.Count}" : "") +
                       (dups > 0 ? $", повторов {dups} (пропущу)" : ", повторов нет") +
                       (bad > 0 ? $", не читаются: {bad}" : ""));
            if (opt.ShopFilter && shopNames.Count > 1 &&
                opt.RegionMode is not (RegionMode.PerRegion or RegionMode.RegionsMerged or RegionMode.PerShop))
            {
                int need = (int)Math.Ceiling(shopNames.Count * opt.ShopPercent / 100m);
                log.Report($"Фильтр по магазинам: товар должен быть минимум в {need} из {shopNames.Count} магазинов ({opt.ShopPercent}%).");
            }
        }
    }

    // ======================= Вспомогательные функции =======================

    /// <summary>Определяет формат по содержимому: XML (YML) начинается с «&lt;», JSON — с «{» или «[».</summary>
    public static SourceFormat DetectFormat(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        var buf = new byte[512];
        int n = fs.Read(buf, 0, buf.Length);

        // UTF-16 BOM: так бывают только XML-файлы (JSON читается как UTF-8)
        if (n >= 2 && ((buf[0] == 0xFF && buf[1] == 0xFE) || (buf[0] == 0xFE && buf[1] == 0xFF)))
            return SourceFormat.Xml;

        int i = (n >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF) ? 3 : 0;   // UTF-8 BOM
        for (; i < n; i++)
        {
            byte b = buf[i];
            if (b == (byte)' ' || b == (byte)'\t' || b == (byte)'\r' || b == (byte)'\n') continue;
            return b == (byte)'<' ? SourceFormat.Xml : SourceFormat.Json;
        }
        return SourceFormat.Json;
    }

    /// <summary>
    /// Название магазина из YML (shop/name, иначе shop/company) — для имени папки результата.
    /// Возвращает null, если файл не YML или названия в нём нет.
    /// </summary>
    public static string? TryReadShopName(string path)
    {
        try
        {
            if (!File.Exists(path) || DetectFormat(path) != SourceFormat.Xml) return null;

            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = XmlReader.Create(fs, CreateXmlSettings());
            string? name = null, company = null;

            reader.MoveToContent();
            while (!reader.EOF)
            {
                // yml_catalog (0) > shop (1) > name / company / offers (2)
                if (reader.NodeType == XmlNodeType.Element && reader.Depth == 2)
                {
                    if (reader.Name == "offers") break;                  // дальше идут только товары
                    if (reader.Name == "name") { name = reader.ReadElementContentAsString(); continue; }
                    if (reader.Name == "company") { company = reader.ReadElementContentAsString(); continue; }
                }
                reader.Read();
            }

            var result = !string.IsNullOrWhiteSpace(name) ? name : company;
            return string.IsNullOrWhiteSpace(result) ? null : result.Trim();
        }
        catch
        {
            return null;
        }
    }

    private static XmlReaderSettings CreateXmlSettings() => new()
    {
        DtdProcessing = DtdProcessing.Ignore,   // в фидах есть <!DOCTYPE ... SYSTEM "shops.dtd">: DTD не загружаем
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        IgnoreProcessingInstructions = true,
        CheckCharacters = false                 // «грязные» управляющие символы в фидах не должны ронять чтение
    };

    // Цена из YML: "195.9", "1 299,90", "1,299.50". Если есть и запятая, и точка, десятичный разделитель — правый.
    private static decimal? ParseYmlPrice(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;

        var sb = new StringBuilder(text.Length);
        foreach (char ch in text)
            if (!char.IsWhiteSpace(ch)) sb.Append(ch);
        string t = sb.ToString();

        int comma = t.LastIndexOf(','), dot = t.LastIndexOf('.');
        if (comma >= 0 && dot >= 0)
        {
            char dec = comma > dot ? ',' : '.';
            char thousands = dec == ',' ? '.' : ',';
            t = t.Replace(thousands.ToString(), "").Replace(dec, '.');
        }
        else if (comma >= 0)
        {
            t = t.Replace(',', '.');
        }

        return decimal.TryParse(t, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign,
                                CultureInfo.InvariantCulture, out var d) ? d : null;
    }

    // Ссылка на картинку: "//cdn..." -> https, относительные достраиваем по адресу магазина, иначе пропускаем
    private static string? ResolvePictureUrl(string? pic, string? shopUrl)
    {
        pic = pic?.Trim();
        if (string.IsNullOrEmpty(pic)) return null;
        if (pic.StartsWith("//")) pic = "https:" + pic;

        if (Uri.TryCreate(pic, UriKind.Absolute, out var abs))
            return abs.Scheme == Uri.UriSchemeHttp || abs.Scheme == Uri.UriSchemeHttps ? pic : null;

        if (!string.IsNullOrEmpty(shopUrl) && Uri.TryCreate(shopUrl, UriKind.Absolute, out var baseUri) &&
            Uri.TryCreate(baseUri, pic, out var combined))
            return combined.AbsoluteUri;

        return null;
    }

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

    private static async Task DownloadJsonAsync(string url, string path, IProgress<string> log, CancellationToken ct, bool quiet = false)
    {
        if (!quiet) log.Report("Скачиваю JSON по ссылке...");
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
                    if (!quiet) log.Report($"JSON: скачано {total / 1024.0 / 1024.0:F0} МБ...");
                }
            }
            if (!quiet) log.Report($"JSON скачан: {total / 1024.0 / 1024.0:F1} МБ");
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
                first.OldPrice ??= g.OldPrice;
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

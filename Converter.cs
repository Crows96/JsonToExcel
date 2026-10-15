using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using ClosedXML.Excel;

namespace JsonToExcel;

public enum PriceMode { Min, Max, Average, Median }

public static class Converter
{
    private sealed class Group
    {
        public string Description = "";
        public string? ImageUrl;
        public List<decimal> Prices = new();
    }

    private sealed record Row(string Description, string? ImageFile, decimal Price);

    private sealed class Part
    {
        public List<Row> Rows = new();
        public HashSet<string> Images = new(StringComparer.OrdinalIgnoreCase);
        public long Bytes;
    }

    public static bool IsUrl(string s) =>
        Uri.TryCreate(s?.Trim(), UriKind.Absolute, out var u) &&
        (u.Scheme == Uri.UriSchemeHttp || u.Scheme == Uri.UriSchemeHttps);

    /// <summary>jsonSource — путь к файлу или ссылка. Возвращает путь к папке с результатом.</summary>
    public static async Task<string> RunAsync(
        string jsonSource, string outDir, decimal maxSpreadPercent, PriceMode priceMode,
        int maxPartMb, bool makeZip, int maxProducts, bool mergeSame,
        IProgress<string> log, IProgress<int> percent, CancellationToken ct)
    {
        var createdParts = new List<string>();   // папки частей, созданные в этом запуске
        try
        {
            return await RunCoreAsync(jsonSource, outDir, maxSpreadPercent, priceMode,
                maxPartMb, makeZip, maxProducts, mergeSame, log, percent, ct, createdParts);
        }
        catch (OperationCanceledException)
        {
            // Пользователь нажал «Стоп»: удаляем всё временное и незаконченное
            CleanupAfterCancel(outDir, createdParts);
            log.Report("Остановлено. Временные файлы удалены.");
            throw;
        }
    }

    private static void CleanupAfterCancel(string outDir, List<string> createdParts)
    {
        foreach (var p in createdParts)
        {
            try { if (Directory.Exists(p)) Directory.Delete(p, true); } catch { }
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

    private static async Task<string> RunCoreAsync(
        string jsonSource, string outDir, decimal maxSpreadPercent, PriceMode priceMode,
        int maxPartMb, bool makeZip, int maxProducts, bool mergeSame,
        IProgress<string> log, IProgress<int> percent, CancellationToken ct, List<string> createdParts)
    {
        Directory.CreateDirectory(outDir);
        string cacheDir = Path.Combine(outDir, "_cache");   // временная папка для скачанных картинок
        Directory.CreateDirectory(cacheDir);

        // ---------- 0. Если указана ссылка — сначала скачиваем JSON ----------
        string jsonPath = jsonSource.Trim();
        string? tempJson = null;
        if (IsUrl(jsonPath))
        {
            tempJson = Path.Combine(outDir, "_source.json");
            await DownloadJsonAsync(jsonPath, tempJson, log, ct);
            jsonPath = tempJson;
        }

        void Cleanup(bool keepCache)
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

        // ---------- 1. Читаем JSON и группируем записи по товару ----------
        log.Report("Читаю JSON...");
        var (groups, offersCount) = ReadGroups(jsonPath, ct);
        ct.ThrowIfCancellationRequested();
        log.Report($"Записей в JSON: {offersCount}, уникальных товаров (по коду): {groups.Count}");

        if (mergeSame)
        {
            int before = groups.Count;
            groups = MergeSameProducts(groups);
            if (groups.Count < before)
                log.Report($"Объединено товаров с одинаковым названием и картинкой: {before - groups.Count} " +
                           $"(осталось {groups.Count})");
        }

        // ---------- 2. Отбор по разбросу цен ----------
        decimal limit = maxSpreadPercent / 100m;
        var selected = new List<(Group Group, decimal Price)>();
        int dropped = 0, noPrice = 0, keptSingle = 0, keptEqual = 0, keptSimilar = 0;

        foreach (var g in groups)
        {
            if (g.Prices.Count == 0) { noPrice++; continue; }

            decimal min = g.Prices.Min();
            decimal max = g.Prices.Max();
            decimal spread = (max - min) / min;          // min > 0 гарантировано при чтении

            if (spread > limit) { dropped++; continue; }

            if (g.Prices.Count == 1) keptSingle++;
            else if (max == min) keptEqual++;
            else keptSimilar++;

            selected.Add((g, PickPrice(g.Prices, priceMode)));
        }

        log.Report($"Порог разброса цен: {maxSpreadPercent}%");
        log.Report($"Взято товаров: {selected.Count} " +
                   $"(цена одинаковая: {keptEqual}, разница в пределах порога: {keptSimilar}, " +
                   $"есть только в одном регионе: {keptSingle})");
        log.Report($"Пропущено из-за разброса цен: {dropped}");
        if (noPrice > 0) log.Report($"Пропущено без корректной цены: {noPrice}");

        if (selected.Count == 0)
        {
            log.Report("Нет товаров для выгрузки.");
            Cleanup(false);
            percent.Report(100);
            return outDir;
        }

        // ---------- 3–4. Картинки: имена файлов и скачивание (во временную папку) ----------
        var urlToName = new Dictionary<string, string>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var failed = new ConcurrentBag<string>();
        int noImageUrl = 0, imageFailed = 0;

        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0");

        // Картинка считается готовой, если файл скачан и не пустой
        bool HasImage(string url)
        {
            if (!urlToName.TryGetValue(url, out var n)) return false;
            var fi = new FileInfo(Path.Combine(cacheDir, n));
            return fi.Exists && fi.Length > 0;
        }

        // Назначает имена файлов и скачивает те картинки, которые ещё не обрабатывались
        async Task EnsureImagesAsync(IEnumerable<string?> urls, Action<int, int>? onProgress)
        {
            var todo = new Dictionary<string, string>();
            foreach (var url in urls.Where(u => !string.IsNullOrWhiteSpace(u)).Distinct())
            {
                if (urlToName.ContainsKey(url!)) continue;      // уже обработана раньше

                string name = MakeFileName(url!);
                string unique = name;
                int k = 1;
                while (!used.Add(unique))
                    unique = $"{Path.GetFileNameWithoutExtension(name)}_{k++}{Path.GetExtension(name)}";
                urlToName[url!] = unique;
                todo[url!] = unique;
            }
            if (todo.Count == 0) return;
            if (onProgress != null) log.Report($"Картинок к скачиванию: {todo.Count}");

            int done = 0;
            int total = todo.Count;
            await Parallel.ForEachAsync(todo, new ParallelOptions { MaxDegreeOfParallelism = 16, CancellationToken = ct },
                async (pair, token) =>
                {
                    string path = Path.Combine(cacheDir, pair.Value);
                    if (!(File.Exists(path) && new FileInfo(path).Length > 0))   // уже скачано — пропускаем
                    {
                        bool ok = false;
                        for (int attempt = 1; attempt <= 3 && !ok; attempt++)
                        {
                            try
                            {
                                var bytes = await http.GetByteArrayAsync(pair.Key, token);
                                await File.WriteAllBytesAsync(path, bytes, token);
                                ok = true;
                            }
                            catch (OperationCanceledException) when (token.IsCancellationRequested)
                            {
                                throw;                       // «Стоп» — не считаем ошибкой скачивания
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

        if (maxProducts > 0)
        {
            // Лимит: берём товары по порядку из JSON и качаем картинки партиями, пока не наберём нужное число.
            // Товары, у которых картинка не скачалась, заменяются следующими.
            log.Report($"Лимит: не больше {maxProducts} товаров");
            var accepted = new List<(Group Group, decimal Price)>();
            int idx = 0;
            while (accepted.Count < maxProducts && idx < selected.Count)
            {
                ct.ThrowIfCancellationRequested();
                int need = maxProducts - accepted.Count;
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
                percent.Report(accepted.Count * 80 / maxProducts);
                log.Report($"Подготовлено товаров: {accepted.Count}/{maxProducts}");
            }

            if (accepted.Count < maxProducts)
                log.Report($"Подходящих товаров меньше лимита: взято {accepted.Count} из {maxProducts}.");
            selected = accepted;
        }
        else
        {
            await EnsureImagesAsync(selected.Select(s => s.Group.ImageUrl), (d, t) =>
            {
                percent.Report(d * 80 / t);                      // 0–80% — скачивание
                if (d % 200 == 0 || d == t) log.Report($"Скачано: {d}/{t}");
            });
        }

        // ---------- 5. Делим товары на части по размеру картинок ----------
        var sizes = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in urlToName.Values)
        {
            var fi = new FileInfo(Path.Combine(cacheDir, name));
            if (fi.Exists && fi.Length > 0) sizes[name] = fi.Length;
        }

        long maxBytes = (long)maxPartMb * 1024 * 1024;
        var parts = new List<Part>();
        var cur = new Part();
        foreach (var (g, price) in selected)
        {
            string? file = null;
            if (g.ImageUrl != null && urlToName.TryGetValue(g.ImageUrl, out var n) && sizes.ContainsKey(n))
                file = n;

            // Нет картинки (ссылки нет или скачать не удалось) — товар в Excel не добавляем
            if (file == null)
            {
                if (string.IsNullOrWhiteSpace(g.ImageUrl)) noImageUrl++; else imageFailed++;
                continue;
            }

            // Картинка, которая уже есть в текущей части, место не занимает
            long add = (file != null && !cur.Images.Contains(file)) ? sizes[file] : 0;
            if (cur.Rows.Count > 0 && cur.Bytes + add > maxBytes)
            {
                parts.Add(cur);
                cur = new Part();
            }

            cur.Rows.Add(new Row(g.Description, file, price));
            if (file != null && cur.Images.Add(file)) cur.Bytes += sizes[file];
        }
        if (cur.Rows.Count > 0) parts.Add(cur);

        if (noImageUrl + imageFailed > 0)
            log.Report($"Не добавлено из-за картинок: {noImageUrl + imageFailed} " +
                       $"(нет ссылки на картинку: {noImageUrl}, не удалось скачать: {imageFailed})");
        if (parts.Count == 0)
            log.Report("Ни один товар не попал в Excel: ни для одного не удалось получить картинку.");
        log.Report($"Частей: {parts.Count} (не больше {maxPartMb} МБ картинок в каждой)");

        // Удаляем части от предыдущего запуска, чтобы не оставалось лишнего
        foreach (var old in Directory.GetDirectories(outDir, "part_*"))
            Directory.Delete(old, true);

        // ---------- 6. Пишем части: Excel + картинки (папка или ZIP) ----------
        for (int i = 0; i < parts.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var part = parts[i];
            string partDir = Path.Combine(outDir, $"part_{i + 1:D3}");
            Directory.CreateDirectory(partDir);
            createdParts.Add(partDir);

            // Excel
            using (var wb = new XLWorkbook())
            {
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
                    // date_start и date_end намеренно оставлены пустыми
                    row++;
                }
                ws.Columns().AdjustToContents();
                wb.SaveAs(Path.Combine(partDir, "catalog.xlsx"));
            }

            // Картинки
            if (makeZip)
            {
                string zipPath = Path.Combine(partDir, "images.zip");
                using var zs = new FileStream(zipPath, FileMode.Create);
                using var za = new ZipArchive(zs, ZipArchiveMode.Create);
                foreach (var img in part.Images)
                {
                    ct.ThrowIfCancellationRequested();
                    za.CreateEntryFromFile(Path.Combine(cacheDir, img), img, CompressionLevel.Fastest);
                }
            }
            else
            {
                string imgDir = Path.Combine(partDir, "images");
                Directory.CreateDirectory(imgDir);
                foreach (var img in part.Images)
                {
                    ct.ThrowIfCancellationRequested();
                    File.Copy(Path.Combine(cacheDir, img), Path.Combine(imgDir, img), true);
                }
            }

            log.Report($"part_{i + 1:D3}: товаров {part.Rows.Count}, картинок {part.Images.Count}, " +
                       $"{part.Bytes / 1024.0 / 1024.0:F1} МБ");
            percent.Report(80 + (i + 1) * 19 / parts.Count);   // 80–99% — упаковка
        }

        // Временную папку с картинками удаляем после упаковки. Если часть картинок не скачалась,
        // оставляем её: при повторном запуске докачаются только недостающие.
        Cleanup(keepCache: !failed.IsEmpty);

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

    // Весь DOM живёт только внутри этого метода и освобождается до скачивания картинок.
    private static (List<Group> Groups, int OffersCount) ReadGroups(string jsonPath, CancellationToken ct)
    {
        var list = new List<Group>();
        var index = new Dictionary<string, Group>();
        int count = 0;

        using var fs = File.OpenRead(jsonPath);
        using var doc = JsonDocument.Parse(fs);

        if (!doc.RootElement.TryGetProperty("offers", out var offers) || offers.ValueKind != JsonValueKind.Array)
            throw new Exception("В JSON нет массива \"offers\".");

        foreach (var o in offers.EnumerateArray())
        {
            count++;
            if ((count & 1023) == 0) ct.ThrowIfCancellationRequested();
            string description = GetString(o, "description") ?? "";
            string id = GetString(o, "id") ?? "";

            // id вида "<код товара>_<регион/магазин>" -> ключ товара = часть до первого "_".
            // Если "_" нет (как в palich), каждая запись — отдельный товар.
            string key;
            int us = id.IndexOf('_');
            if (us > 0) key = id[..us];
            else key = id.Length > 0 ? id : description;

            if (!index.TryGetValue(key, out var g))
            {
                g = new Group();
                index[key] = g;
                list.Add(g);
            }

            if (g.Description.Length == 0) g.Description = description;
            if (string.IsNullOrWhiteSpace(g.ImageUrl)) g.ImageUrl = GetString(o, "image");

            var price = GetDecimal(o, "price_new");
            if (price.HasValue && price.Value > 0) g.Prices.Add(price.Value);
        }

        return (list, count);
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
                first.Prices.AddRange(g.Prices);
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

    private static decimal? GetDecimal(JsonElement e, string key)
    {
        if (!e.TryGetProperty(key, out var v)) return null;
        if (v.ValueKind == JsonValueKind.Number) return v.TryGetDecimal(out var n) ? n : null;
        if (v.ValueKind == JsonValueKind.String &&
            decimal.TryParse(v.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
        return null;
    }

    private static string MakeFileName(string url)
    {
        string name;
        try { name = Path.GetFileName(new Uri(url).AbsolutePath); }
        catch { name = ""; }
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        if (string.IsNullOrWhiteSpace(name)) name = Math.Abs(url.GetHashCode()).ToString();
        if (!Path.HasExtension(name)) name += ".jpg";
        return name;
    }
}

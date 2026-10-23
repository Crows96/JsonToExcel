using System.Text;
using System.Text.RegularExpressions;

namespace JsonToExcel;

/// <summary>
/// Справочник городов и регионов России (встроенный ресурс ru_geo.tsv). Нужен, чтобы по адресу магазина
/// или по названию из target_regions получить официальный регион (субъект РФ): область, край, республику.
/// Если ресурса нет, Loaded = false, и программа работает как раньше (регион — город из адреса).
/// </summary>
public static class RuGeo
{
    private sealed record RegionInfo(string Code, string FullName);

    private static readonly Dictionary<string, RegionInfo> regionByKey = new();
    private static readonly Dictionary<string, RegionInfo> regionByCode = new();
    private static readonly Dictionary<string, List<(string Code, int Population)>> cityByName = new();

    public static bool Loaded { get; private set; }

    // Короткие названия, которых нет в справочнике
    private static readonly (string Alias, string Code)[] Aliases =
    {
        ("хмао", "86"), ("янао", "89"), ("мо", "50"), ("спб", "78"), ("питер", "78"), ("крым", "91")
    };

    // Территории Москвы, которых нет среди городов справочника (Новая Москва, Зеленоград и т. д.)
    private static readonly HashSet<string> MoscowTerritories = new()
    {
        "зеленоград", "московский", "сосенское", "щербинка", "мосрентген", "внуково", "десеновское",
        "кокошкино", "филимонковское", "вороновское", "марушкинское", "рязановское"
    };

    static RuGeo()
    {
        try { Load(); }
        catch { Loaded = false; }
    }

    private static void Load()
    {
        var asm = typeof(RuGeo).Assembly;
        string? res = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("ru_geo.tsv", StringComparison.OrdinalIgnoreCase));
        if (res == null) return;

        using var stream = asm.GetManifestResourceStream(res)!;
        using var reader = new StreamReader(stream, Encoding.UTF8);
        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            var p = line.Split('\t');
            if (p.Length >= 5 && p[0] == "R")
            {
                var info = new RegionInfo(p[1], p[2]);
                regionByCode[p[1]] = info;
                foreach (var raw in new[] { p[2], p[3], p[4] })
                {
                    string key = StripKey(raw);
                    if (key.Length > 0) regionByKey.TryAdd(key, info);
                }
            }
            else if (p.Length >= 4 && p[0] == "C")
            {
                string key = Norm(p[1]);
                if (!cityByName.TryGetValue(key, out var list)) cityByName[key] = list = new();
                int.TryParse(p[3], out int pop);
                list.Add((p[2], pop));
            }
        }

        foreach (var (alias, code) in Aliases)
            if (regionByCode.TryGetValue(code, out var info)) regionByKey.TryAdd(alias, info);

        Loaded = regionByCode.Count > 0;
    }

    // ---------- Нормализация текста ----------
    private static string Norm(string s)
    {
        s = s.ToLowerInvariant().Replace('ё', 'е');
        var sb = new StringBuilder(s.Length);
        foreach (char ch in s) sb.Append(char.IsLetterOrDigit(ch) || ch == '-' ? ch : ' ');
        return Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
    }

    private static readonly HashSet<string> RegionWords = new()
    {
        "область", "обл", "край", "республика", "респ", "автономный", "автономная", "ао", "округ", "г", "город", "-"
    };

    // Название региона без слов «область», «край», «республика» и т. п.
    private static string StripKey(string s) =>
        string.Join(' ', Norm(s).Split(' ', StringSplitOptions.RemoveEmptyEntries).Where(w => !RegionWords.Contains(w)));

    // ---------- Регион по названию ----------
    /// <summary>Официальное название региона по строке вроде «Московская обл.», «Респ Татарстан», «Якутия». null — не найден.</summary>
    public static string? RegionName(string text)
    {
        if (!Loaded) return null;
        string key = StripKey(text);
        return key.Length > 0 && regionByKey.TryGetValue(key, out var info) ? info.FullName : null;
    }

    // Самый крупный город с таким названием (если в адресе нет региона, выбирать больше не из чего)
    private static string? RegionOfCityName(string city, string? regionCode = null)
    {
        if (!cityByName.TryGetValue(Norm(city), out var list)) return null;
        var pool = regionCode == null ? list : list.Where(x => x.Code == regionCode).ToList();
        if (pool.Count == 0) return null;
        string code = pool.OrderByDescending(x => x.Population).First().Code;
        return regionByCode.TryGetValue(code, out var info) ? info.FullName : null;
    }

    /// <summary>
    /// Нормализация значения из target_regions: сначала регион по названию; если включено citiesToo —
    /// затем город (его регион). null — не распознано, значение остаётся как есть.
    /// </summary>
    public static string? NormalizeRegion(string text, bool citiesToo)
    {
        string? r = RegionName(text);
        if (r != null) return r;
        return citiesToo && Loaded ? RegionOfCityName(StripPrefix(text)) : null;
    }

    private static string StripPrefix(string s)
    {
        var m = CityMarker.Match(s.Trim());
        return m.Success ? m.Groups["n"].Value : s;
    }

    // ---------- Регион по адресу ----------
    private static readonly Regex CityMarker = new(
        @"^(?:г|город)(?:\.\s*|\s+)(?<n>\p{L}.*)$|^(?<n>\p{L}.*?)\s+(?:г|гор)\.?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex OtherSettlement = new(
        @"^(?:пгт|рп|р\.п|пос|посёлок|поселок|п|с|село|ст-ца|станица|х|хутор|аул|дер|деревня|мкр)(?:\.\s*|\s+)(?<n>\p{L}.*)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex NotCityPart = new(
        @"^(?:россия|рф|russia|\d.*)$|^(?:ул|улица|пр|пр-т|просп|проспект|пер|переулок|б-р|бульвар|ш|шоссе|наб|набережная|пл|площадь|тракт|проезд|аллея|линия|мкр|микрорайон|тц|трц|тк|д|дом|к|корп|стр|вл|кв|пом|литер|лит)(?:\.|\s|$)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Регион (субъект РФ) по адресу магазина: «г. Химки, ул. Мира, 3» -> «Московская область».
    /// Сначала ищется регион в самом адресе, затем город в справочнике; если одноимённых городов несколько,
    /// берётся самый крупный. Посёлки (п., р.п., пгт) в справочнике городов не ищутся. null — определить не удалось.
    /// </summary>
    public static string? RegionOfAddress(string address)
    {
        if (!Loaded) return null;
        var parts = address.Split(',').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        if (parts.Count == 0) return null;

        // 1. Регион указан прямо в адресе («Московская обл.», «Респ Татарстан»)
        string? regionCode = null;
        foreach (var p in parts)
        {
            string key = StripKey(p);
            if (key.Length > 0 && regionByKey.TryGetValue(key, out var info)) { regionCode = info.Code; break; }
        }

        // 2. Населённый пункт
        string? city = null, other = null;
        foreach (var p in parts)
        {
            var m = CityMarker.Match(p);
            if (m.Success) { city = m.Groups["n"].Value.Trim(' ', '.'); break; }
            m = OtherSettlement.Match(p);
            if (m.Success && other == null) other = m.Groups["n"].Value.Trim(' ', '.');
        }
        if (city == null && other == null)
        {
            // Без пометки («Москва, ул. Ленина, 5»): первая часть, не похожая на страну, индекс или улицу
            foreach (var p in parts)
            {
                if (NotCityPart.IsMatch(p)) continue;
                string key = StripKey(p);
                if (regionByKey.ContainsKey(key) && !cityByName.ContainsKey(Norm(p))) continue;   // это название региона
                city = p.Trim(' ', '.');
                break;
            }
        }

        if (regionCode != null)
            return regionByCode[regionCode].FullName;

        if (city != null)
        {
            string? r = RegionOfCityName(city);
            if (r != null) return r;
        }

        // 3. Москва: города Новой Москвы и Зеленоград в справочнике городов отсутствуют
        string? name = city ?? other;
        if (name != null && MoscowTerritories.Contains(Norm(name)) && regionByCode.TryGetValue("77", out var moscow))
            return moscow.FullName;

        return null;
    }
}

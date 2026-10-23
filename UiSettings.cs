using System.Text.Json;

namespace JsonToExcel;

/// <summary>Настройки окна, которые запоминаются между запусками и сохраняются в наборах.</summary>
public sealed class UiSettings
{
    public decimal Spread { get; set; } = 0;
    public int PriceMode { get; set; } = 1;
    public int PartMb { get; set; } = 100;
    public bool Zip { get; set; } = true;
    public int Limit { get; set; }
    public bool Merge { get; set; } = true;
    public bool PromoOnly { get; set; }
    public bool ShopFilter { get; set; } = true;
    public decimal ShopPercent { get; set; } = 95;
    public int GeoMode { get; set; }
    public bool RegionBySubject { get; set; }
    public bool Combine { get; set; }
    public bool Economy { get; set; }
    public int Threads { get; set; } = 16;
    public bool RecodeEnabled { get; set; } = true;
    public int RecodeFormat { get; set; }
    public int RecodeMaxSide { get; set; } = 1000;
    public int RecodeQuality { get; set; } = 85;
}

/// <summary>Файл настроек: последние использованные настройки и именованные наборы.</summary>
public sealed class SettingsFile
{
    public UiSettings Last { get; set; } = new();
    public Dictionary<string, UiSettings> Presets { get; set; } = new();
}

public static class SettingsStore
{
    private static string FilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JsonToExcel", "settings.json");

    /// <summary>Читает настройки; если файла нет или он повреждён, возвращает значения по умолчанию.</summary>
    public static SettingsFile Load(out bool existed)
    {
        existed = false;
        try
        {
            if (!File.Exists(FilePath)) return new SettingsFile();
            var file = JsonSerializer.Deserialize<SettingsFile>(File.ReadAllText(FilePath));
            if (file == null) return new SettingsFile();
            existed = true;
            file.Last ??= new UiSettings();
            file.Presets ??= new Dictionary<string, UiSettings>();
            return file;
        }
        catch
        {
            return new SettingsFile();
        }
    }

    /// <summary>Сохраняет настройки; ошибки записи (например, нет прав) не должны мешать работе.</summary>
    public static void Save(SettingsFile file)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(file, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }
}

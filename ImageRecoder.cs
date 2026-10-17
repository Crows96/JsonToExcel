using SkiaSharp;

namespace JsonToExcel;

public sealed record RecodeResult(byte[] Data, int SrcWidth, int SrcHeight, int Width, int Height);

/// <summary>
/// Перекодирование картинки через SkiaSharp (нативная библиотека Skia): быстрее и экономнее по памяти,
/// чем System.Drawing, потокобезопасно. Уменьшение и заливка фона делаются за один проход.
/// </summary>
public static class ImageRecoder
{
    public static RecodeResult Recode(byte[] source, RecodeOptions o)
    {
        using var src = SKBitmap.Decode(source);
        if (src == null || src.Width <= 0 || src.Height <= 0)
            throw new InvalidDataException("не удалось прочитать картинку (возможно, это не изображение)");

        // Размер: только уменьшаем, меньшие картинки не растягиваем
        int w = src.Width, h = src.Height;
        int longest = Math.Max(w, h);
        if (o.MaxSide > 0 && longest > o.MaxSide)
        {
            double k = (double)o.MaxSide / longest;
            w = Math.Max(1, (int)Math.Round(w * k));
            h = Math.Max(1, (int)Math.Round(h * k));
        }

        // Для JPEG и WebP прозрачность заливаем белым, для PNG сохраняем
        bool keepAlpha = o.Format == RecodeFormat.Png;
        var info = new SKImageInfo(w, h, SKImageInfo.PlatformColorType,
                                   keepAlpha ? SKAlphaType.Premul : SKAlphaType.Opaque);

        using var dst = new SKBitmap(info);
        using (var canvas = new SKCanvas(dst))
        using (var paint = new SKPaint { FilterQuality = SKFilterQuality.High, IsAntialias = true })
        {
            canvas.Clear(keepAlpha ? SKColors.Transparent : SKColors.White);
            canvas.DrawBitmap(src, new SKRect(0, 0, w, h), paint);
        }

        var format = o.Format switch
        {
            RecodeFormat.Jpeg => SKEncodedImageFormat.Jpeg,
            RecodeFormat.WebP => SKEncodedImageFormat.Webp,
            _ => SKEncodedImageFormat.Png
        };

        using var image = SKImage.FromBitmap(dst);
        using var data = image.Encode(format, Math.Clamp(o.Quality, 1, 100));
        if (data == null)
            throw new InvalidDataException("не удалось закодировать картинку");

        return new RecodeResult(data.ToArray(), src.Width, src.Height, w, h);
    }
}

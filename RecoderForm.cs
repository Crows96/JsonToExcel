namespace JsonToExcel;

/// <summary>Окно настроек сжатия картинок с проверкой на одном файле.</summary>
public class RecoderForm : Form
{
    private readonly CheckBox chkEnabled = new()
    {
        Text = "Сжимать картинки при скачивании", AutoSize = true, Margin = new Padding(3, 3, 3, 10)
    };
    private readonly ComboBox cmbFormat = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 380 };
    private readonly NumericUpDown numSide = new() { Minimum = 0, Maximum = 10000, Increment = 50, Width = 80 };
    private readonly NumericUpDown numQuality = new() { Minimum = 1, Maximum = 100, Width = 80 };
    private readonly Label lblHint = new()
    {
        AutoSize = true, MaximumSize = new Size(540, 0), ForeColor = SystemColors.GrayText,
        Margin = new Padding(3, 10, 3, 10),
        Text = "Как правило, PNG на 1–3 МБ после перевода в JPEG (до 1000 px, качество 85) становится в разы легче, " +
               "поэтому каталог делится на меньшее число частей. Прозрачный фон при сохранении в JPEG и WebP " +
               "заливается белым. Подберите настройки на одной картинке кнопкой ниже."
    };
    private readonly Button btnTest = new() { Text = "Проверить на картинке...", AutoSize = true };
    private readonly Button btnSave = new() { Text = "Сохранить результат...", AutoSize = true, Enabled = false };
    private readonly Label lblResult = new()
    {
        AutoSize = true, MaximumSize = new Size(540, 0), Margin = new Padding(3, 8, 3, 8)
    };
    private readonly Button btnOk = new() { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
    private readonly Button btnCancel = new() { Text = "Отмена", DialogResult = DialogResult.Cancel, Width = 90 };

    private byte[]? lastResult;
    private string lastExtension = ".jpg";

    /// <summary>Настройки, выбранные пользователем (актуальны после закрытия окна кнопкой OK).</summary>
    public RecodeOptions Result { get; private set; }

    public RecoderForm(RecodeOptions current)
    {
        Result = current;

        Text = "Сжатие картинок";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(580, 380);
        AcceptButton = btnOk;
        CancelButton = btnCancel;

        // Порядок пунктов совпадает с enum RecodeFormat
        cmbFormat.Items.AddRange(new object[]
        {
            "JPEG — самый универсальный и лёгкий, без прозрачности",
            "WebP — ещё легче, но подходит не везде",
            "PNG — без потерь, только уменьшение размера"
        });

        // Если сжатие было выключено, показываем рекомендуемые значения
        var shown = current.Enabled ? current : RecodeOptions.Default;
        chkEnabled.Checked = current.Enabled;
        cmbFormat.SelectedIndex = (int)shown.Format;
        numSide.Value = Math.Clamp(shown.MaxSide, 0, 10000);
        numQuality.Value = Math.Clamp(shown.Quality, 1, 100);

        var root = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false,
            Padding = new Padding(14)
        };
        root.Controls.Add(chkEnabled);
        root.Controls.Add(Row("Формат:", cmbFormat));
        root.Controls.Add(Row("Макс. сторона, px (0 — не уменьшать):", numSide));
        root.Controls.Add(Row("Качество, 1–100 (для JPEG и WebP):", numQuality));
        root.Controls.Add(lblHint);

        var testRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        testRow.Controls.Add(btnTest);
        testRow.Controls.Add(btnSave);
        root.Controls.Add(testRow);
        root.Controls.Add(lblResult);

        var okRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(3, 10, 3, 3) };
        okRow.Controls.Add(btnOk);
        okRow.Controls.Add(btnCancel);
        root.Controls.Add(okRow);

        Controls.Add(root);

        chkEnabled.CheckedChanged += (_, _) => UpdateControls();
        cmbFormat.SelectedIndexChanged += (_, _) => UpdateControls();
        btnTest.Click += (_, _) => TestOnFile();
        btnSave.Click += (_, _) => SaveResult();

        // Результат берём в момент закрытия окна кнопкой OK
        FormClosing += (_, _) =>
        {
            if (DialogResult == DialogResult.OK) Result = Build();
        };

        UpdateControls();
    }

    private static FlowLayoutPanel Row(string text, Control control)
    {
        var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        row.Controls.Add(new Label { Text = text, AutoSize = true, Margin = new Padding(0, 6, 8, 0) });
        row.Controls.Add(control);
        return row;
    }

    private RecodeOptions Build() =>
        new(chkEnabled.Checked, (RecodeFormat)Math.Max(0, cmbFormat.SelectedIndex), (int)numSide.Value, (int)numQuality.Value);

    private void UpdateControls()
    {
        bool on = chkEnabled.Checked;
        cmbFormat.Enabled = numSide.Enabled = on;
        numQuality.Enabled = on && cmbFormat.SelectedIndex != (int)RecodeFormat.Png;
    }

    private static string FormatSize(long bytes) =>
        bytes >= 1024 * 1024 ? $"{bytes / 1024.0 / 1024.0:F2} МБ" : $"{bytes / 1024.0:F0} КБ";

    private void TestOnFile()
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Картинки|*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif|Все файлы|*.*"
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        try
        {
            byte[] source = File.ReadAllBytes(dlg.FileName);
            var options = Build() with { Enabled = true };

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var r = ImageRecoder.Recode(source, options);
            sw.Stop();

            lastResult = r.Data;
            lastExtension = options.Extension;
            btnSave.Enabled = true;

            double saved = 100.0 * (1 - (double)r.Data.Length / source.Length);
            string sign = saved >= 0 ? "меньше на" : "больше на";
            lblResult.Text =
                $"{Path.GetFileName(dlg.FileName)}: {r.SrcWidth}×{r.SrcHeight} → {r.Width}×{r.Height}\n" +
                $"{FormatSize(source.Length)} → {FormatSize(r.Data.Length)} ({sign} {Math.Abs(saved):F0}%), {sw.ElapsedMilliseconds} мс";
        }
        catch (Exception ex)
        {
            lblResult.Text = "Ошибка: " + ex.Message;
        }
    }

    private void SaveResult()
    {
        if (lastResult == null) return;
        using var dlg = new SaveFileDialog
        {
            FileName = "preview" + lastExtension,
            Filter = "Картинка|*" + lastExtension
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;
        File.WriteAllBytes(dlg.FileName, lastResult);
    }
}

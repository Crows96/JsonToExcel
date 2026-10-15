namespace JsonToExcel;

public class MainForm : Form
{
    private readonly TextBox txtJson = new()
    {
        Dock = DockStyle.Fill, PlaceholderText = "Путь к файлу или ссылка https://..."
    };
    private readonly TextBox txtOut = new() { Dock = DockStyle.Fill };
    private readonly Button btnJson = new() { Text = "Обзор...", AutoSize = true };
    private readonly Button btnOut = new() { Text = "Обзор...", AutoSize = true };
    private readonly NumericUpDown numSpread = new()
    {
        DecimalPlaces = 1, Minimum = 0, Maximum = 100, Increment = 0.5m, Value = 5, Width = 70,
        Margin = new Padding(0, 3, 20, 0)
    };
    private readonly ComboBox cmbPrice = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList, Width = 140, Margin = new Padding(0, 3, 0, 0)
    };
    private readonly NumericUpDown numPartMb = new()
    {
        Minimum = 1, Maximum = 100000, Value = 100, Width = 80, Margin = new Padding(0, 3, 20, 0)
    };
    private readonly CheckBox chkZip = new()
    {
        Text = "Упаковывать картинки в ZIP", Checked = true, AutoSize = true, Margin = new Padding(0, 4, 0, 0)
    };
    private readonly NumericUpDown numLimit = new()
    {
        Minimum = 0, Maximum = 1000000, Value = 0, Width = 80, Margin = new Padding(0, 3, 0, 0)
    };
    private readonly CheckBox chkMerge = new()
    {
        Text = "Считать одним товаром позиции с одинаковым названием и картинкой",
        Checked = true, AutoSize = true, Margin = new Padding(0, 4, 0, 0)
    };
    private readonly Button btnStart = new() { Text = "Старт", Height = 34, Dock = DockStyle.Fill };
    private readonly Button btnStop = new() { Text = "Стоп", Height = 34, Dock = DockStyle.Fill, Enabled = false };
    private CancellationTokenSource? cts;
    private readonly ProgressBar progress = new() { Dock = DockStyle.Fill };
    private readonly TextBox log = new()
    {
        Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill
    };

    public MainForm()
    {
        Text = "JSON → Excel + картинки";
        Width = 760;
        Height = 640;
        StartPosition = FormStartPosition.CenterScreen;

        cmbPrice.Items.AddRange(new object[] { "Минимальная", "Максимальная", "Средняя", "Медиана" });
        cmbPrice.SelectedIndex = 0;

        var settings = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settings.Controls.Add(new Label { Text = "Макс. разброс цен по регионам, %:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settings.Controls.Add(numSpread);
        settings.Controls.Add(new Label { Text = "Цена в Excel:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settings.Controls.Add(cmbPrice);

        var settings2 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settings2.Controls.Add(new Label { Text = "Макс. размер картинок в одной части, МБ:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settings2.Controls.Add(numPartMb);
        settings2.Controls.Add(chkZip);

        var settings3 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settings3.Controls.Add(new Label { Text = "Макс. число товаров (0 — без ограничения):", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settings3.Controls.Add(numLimit);

        var settings4 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settings4.Controls.Add(chkMerge);

        var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        buttons.Controls.Add(btnStart, 0, 0);
        buttons.Controls.Add(btnStop, 1, 0);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 9, Padding = new Padding(10)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        layout.Controls.Add(new Label { Text = "JSON (файл или ссылка):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        layout.Controls.Add(txtJson, 1, 0);
        layout.Controls.Add(btnJson, 2, 0);

        layout.Controls.Add(new Label { Text = "Папка результата:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        layout.Controls.Add(txtOut, 1, 1);
        layout.Controls.Add(btnOut, 2, 1);

        layout.Controls.Add(settings, 0, 2);
        layout.SetColumnSpan(settings, 3);

        layout.Controls.Add(settings2, 0, 3);
        layout.SetColumnSpan(settings2, 3);

        layout.Controls.Add(settings3, 0, 4);
        layout.SetColumnSpan(settings3, 3);

        layout.Controls.Add(settings4, 0, 5);
        layout.SetColumnSpan(settings4, 3);

        layout.Controls.Add(buttons, 0, 6);
        layout.SetColumnSpan(buttons, 3);

        layout.Controls.Add(progress, 0, 7);
        layout.SetColumnSpan(progress, 3);

        layout.Controls.Add(log, 0, 8);
        layout.SetColumnSpan(log, 3);

        Controls.Add(layout);

        btnJson.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "JSON (*.json)|*.json|Все файлы|*.*" };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            txtJson.Text = dlg.FileName;
            ApplySuggestion();   // на случай, если выбран тот же файл, а папку успели поменять
        };

        btnOut.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK) txtOut.Text = dlg.SelectedPath;
        };

        // Любое изменение JSON (выбор файла, вставка ссылки или пути) -> папка результата пересчитывается
        txtJson.TextChanged += (_, _) => ApplySuggestion();

        btnStart.Click += async (_, _) => await StartAsync();

        btnStop.Click += (_, _) =>
        {
            if (cts == null) return;
            btnStop.Enabled = false;
            log.AppendText("Останавливаю... (подождите несколько секунд)" + Environment.NewLine);
            cts.Cancel();
        };
    }

    // Папка результата называется так же, как JSON:
    //  - файл  -> рядом с файлом;
    //  - ссылка -> в "Документах", имя берётся из последней части ссылки.
    private static string? SuggestOut(string source)
    {
        source = source.Trim();

        if (Converter.IsUrl(source))
        {
            string name = "result";
            try
            {
                var u = new Uri(source);
                var seg = u.Segments.Select(s => s.Trim('/')).LastOrDefault(s => s.Length > 0);
                name = !string.IsNullOrEmpty(seg) ? Path.GetFileNameWithoutExtension(seg) : u.Host;
            }
            catch { }
            foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            if (string.IsNullOrWhiteSpace(name)) name = "result";
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), name);
        }

        if (File.Exists(source))
            return Path.Combine(Path.GetDirectoryName(source)!, Path.GetFileNameWithoutExtension(source));

        return null;
    }

    // Папка результата всегда следует за JSON. Если нужна своя папка — выберите её после выбора JSON.
    private void ApplySuggestion()
    {
        var s = SuggestOut(txtJson.Text);
        if (s != null) txtOut.Text = s;
    }

    private void SetBusy(bool busy)
    {
        btnStart.Enabled = btnJson.Enabled = btnOut.Enabled = !busy;
        btnStop.Enabled = busy;
        numSpread.Enabled = cmbPrice.Enabled = numPartMb.Enabled = chkZip.Enabled = numLimit.Enabled = chkMerge.Enabled = !busy;
    }

    private async Task StartAsync()
    {
        string source = txtJson.Text.Trim();
        if (!Converter.IsUrl(source) && !File.Exists(source))
        {
            MessageBox.Show("Укажите существующий JSON-файл или ссылку (http:// или https://).", "Ошибка");
            return;
        }
        if (string.IsNullOrWhiteSpace(txtOut.Text)) ApplySuggestion();
        if (string.IsNullOrWhiteSpace(txtOut.Text))
        {
            MessageBox.Show("Укажите папку для результата.", "Ошибка");
            return;
        }

        SetBusy(true);
        log.Clear();
        progress.Value = 0;

        decimal spread = numSpread.Value;
        var mode = (PriceMode)cmbPrice.SelectedIndex;   // порядок пунктов совпадает с enum PriceMode
        int partMb = (int)numPartMb.Value;
        bool zip = chkZip.Checked;
        int limit = (int)numLimit.Value;
        bool merge = chkMerge.Checked;

        cts = new CancellationTokenSource();
        var token = cts.Token;

        // Progress<T> создан в UI-потоке, поэтому обработчики выполняются тоже в нём
        var logProgress = new Progress<string>(s => log.AppendText(s + Environment.NewLine));
        var percentProgress = new Progress<int>(p => progress.Value = Math.Clamp(p, 0, 100));

        try
        {
            string result = await Task.Run(() =>
                Converter.RunAsync(source, txtOut.Text, spread, mode, partMb, zip, limit, merge, logProgress, percentProgress, token));

            progress.Value = 100;
            MessageBox.Show($"Готово!\n\n{result}", "JSON → Excel");
        }
        catch (OperationCanceledException)
        {
            progress.Value = 0;   // сообщение об остановке уже выведено в лог
        }
        catch (Exception ex)
        {
            log.AppendText("ОШИБКА: " + ex.Message + Environment.NewLine);
            MessageBox.Show(ex.Message, "Ошибка");
        }
        finally
        {
            cts?.Dispose();
            cts = null;
            SetBusy(false);
        }
    }
}

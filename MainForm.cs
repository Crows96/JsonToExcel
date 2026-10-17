namespace JsonToExcel;

public class MainForm : Form
{
    private readonly TextBox txtJson = new()
    {
        Dock = DockStyle.Fill, PlaceholderText = "Путь к файлу или ссылка https://... (Enter — добавить в список)"
    };
    private readonly Button btnJson = new() { Text = "Файлы...", AutoSize = true };
    private readonly Button btnAddLink = new() { Text = "Добавить", AutoSize = true };
    private readonly ListBox lstSources = new()
    {
        Dock = DockStyle.Fill, SelectionMode = SelectionMode.MultiExtended, HorizontalScrollbar = true,
        IntegralHeight = false, AllowDrop = true
    };
    private readonly Button btnRemove = new() { Text = "Убрать", AutoSize = true };
    private readonly Button btnClear = new() { Text = "Очистить", AutoSize = true };
    private readonly TextBox txtOut = new() { Dock = DockStyle.Fill };
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
    private readonly CheckBox chkGeo = new()
    {
        Text = "Фильтр по магазинам (target_shops): товар минимум в", AutoSize = true, Margin = new Padding(0, 4, 4, 0)
    };
    private readonly NumericUpDown numGeo = new()
    {
        Minimum = 1, Maximum = 100, Value = 80, Width = 60, Margin = new Padding(0, 3, 4, 0)
    };
    private readonly ComboBox cmbRegion = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList, Width = 480, Margin = new Padding(0, 3, 0, 0)
    };
    private readonly Label lblRecode = new() { AutoSize = true, Margin = new Padding(0, 6, 10, 0) };
    private readonly Button btnRecode = new() { Text = "Настроить...", AutoSize = true };
    private readonly CheckBox chkEconomy = new()
    {
        Text = "Режим для слабого ПК", AutoSize = true, Margin = new Padding(0, 4, 20, 0)
    };
    private readonly NumericUpDown numThreads = new()
    {
        Minimum = 1, Maximum = 64, Value = 16, Width = 60, Margin = new Padding(0, 3, 0, 0)
    };
    private readonly Button btnStart = new() { Text = "Старт", Height = 34, Dock = DockStyle.Fill };
    private readonly Button btnStop = new() { Text = "Стоп", Height = 34, Dock = DockStyle.Fill, Enabled = false };
    private CancellationTokenSource? cts;
    private readonly ProgressBar progress = new() { Dock = DockStyle.Fill };
    private readonly TextBox log = new()
    {
        Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill
    };

    private RecodeOptions recode = RecodeOptions.Default;
    private bool isBusy;

    public MainForm()
    {
        Text = "JSON → Excel + картинки";
        Width = 800;
        Height = 900;
        StartPosition = FormStartPosition.CenterScreen;

        cmbPrice.Items.AddRange(new object[] { "Минимальная", "Максимальная", "Средняя", "Медиана" });
        cmbPrice.SelectedIndex = 0;

        // Порядок пунктов совпадает с enum RegionMode
        cmbRegion.Items.AddRange(new object[]
        {
            "Не разбивать (один Excel на весь файл)",
            "Один Excel: только товары, которые есть во всех регионах (target_regions)",
            "Отдельный Excel (папка) на каждый регион (target_regions)",
            "Отдельный Excel (папка) на каждый магазин / адрес (target_shops)"
        });
        cmbRegion.SelectedIndex = 0;

        UpdateRecodeLabel();

        var sourceButtons = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        sourceButtons.Controls.Add(btnJson);
        sourceButtons.Controls.Add(btnAddLink);

        var listButtons = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.TopDown };
        listButtons.Controls.Add(btnRemove);
        listButtons.Controls.Add(btnClear);

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

        var settings5 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settings5.Controls.Add(chkGeo);
        settings5.Controls.Add(numGeo);
        settings5.Controls.Add(new Label { Text = "% магазинов", AutoSize = true, Margin = new Padding(0, 6, 0, 0) });

        var settings6 = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settings6.Controls.Add(new Label { Text = "Разбивка по гео:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settings6.Controls.Add(cmbRegion);

        var settingsRecode = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settingsRecode.Controls.Add(new Label { Text = "Сжатие картинок:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settingsRecode.Controls.Add(lblRecode);
        settingsRecode.Controls.Add(btnRecode);

        var settingsEco = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
        settingsEco.Controls.Add(chkEconomy);
        settingsEco.Controls.Add(new Label { Text = "Потоков скачивания:", AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        settingsEco.Controls.Add(numThreads);

        var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        buttons.Controls.Add(btnStart, 0, 0);
        buttons.Controls.Add(btnStop, 1, 0);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 14, Padding = new Padding(10)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));           // 0: добавить JSON
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));      // 1: очередь файлов
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));           // 2: папка результата
        for (int i = 0; i < 8; i++)                                      // 3-10: настройки
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));       // 11: кнопки
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));       // 12: прогресс
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));       // 13: лог

        layout.Controls.Add(new Label { Text = "Добавить JSON:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        layout.Controls.Add(txtJson, 1, 0);
        layout.Controls.Add(sourceButtons, 2, 0);

        layout.Controls.Add(new Label { Text = "Очередь (по очереди):", AutoSize = true, Anchor = AnchorStyles.Top | AnchorStyles.Left, Margin = new Padding(3, 6, 3, 0) }, 0, 1);
        layout.Controls.Add(lstSources, 1, 1);
        layout.Controls.Add(listButtons, 2, 1);

        layout.Controls.Add(new Label { Text = "Папка результата:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        layout.Controls.Add(txtOut, 1, 2);
        layout.Controls.Add(btnOut, 2, 2);

        layout.Controls.Add(settings, 0, 3);
        layout.SetColumnSpan(settings, 3);

        layout.Controls.Add(settings2, 0, 4);
        layout.SetColumnSpan(settings2, 3);

        layout.Controls.Add(settings3, 0, 5);
        layout.SetColumnSpan(settings3, 3);

        layout.Controls.Add(settings4, 0, 6);
        layout.SetColumnSpan(settings4, 3);

        layout.Controls.Add(settings5, 0, 7);
        layout.SetColumnSpan(settings5, 3);

        layout.Controls.Add(settings6, 0, 8);
        layout.SetColumnSpan(settings6, 3);

        layout.Controls.Add(settingsRecode, 0, 9);
        layout.SetColumnSpan(settingsRecode, 3);

        layout.Controls.Add(settingsEco, 0, 10);
        layout.SetColumnSpan(settingsEco, 3);

        layout.Controls.Add(buttons, 0, 11);
        layout.SetColumnSpan(buttons, 3);

        layout.Controls.Add(progress, 0, 12);
        layout.SetColumnSpan(progress, 3);

        layout.Controls.Add(log, 0, 13);
        layout.SetColumnSpan(log, 3);

        Controls.Add(layout);

        // ----- Список JSON -----
        btnJson.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "JSON (*.json)|*.json|Все файлы|*.*", Multiselect = true };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            foreach (var f in dlg.FileNames) AddSource(f);
            ApplySuggestion();
        };

        btnAddLink.Click += (_, _) => AddFromTextBox();

        txtJson.KeyDown += (_, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            e.SuppressKeyPress = true;
            AddFromTextBox();
        };

        btnRemove.Click += (_, _) =>
        {
            foreach (var item in lstSources.SelectedItems.Cast<object>().ToList())
                lstSources.Items.Remove(item);
            ApplySuggestion();
        };

        btnClear.Click += (_, _) => lstSources.Items.Clear();

        // Файлы можно просто перетащить в список
        lstSources.DragEnter += (_, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy;
        };
        lstSources.DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] files) return;
            foreach (var f in files) AddSource(f);
            ApplySuggestion();
        };

        btnOut.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK) txtOut.Text = dlg.SelectedPath;
        };

        // Охват по магазинам и регионам считается по объединённым товарам, поэтому включаем объединение автоматически
        chkGeo.CheckedChanged += (_, _) => { if (chkGeo.Checked) chkMerge.Checked = true; };
        cmbRegion.SelectedIndexChanged += (_, _) =>
        {
            // Для режимов по регионам охват считается по объединённым товарам; для магазинов объединение не обязательно
            if (cmbRegion.SelectedIndex is (int)RegionMode.AllRegions or (int)RegionMode.PerRegion)
                chkMerge.Checked = true;
            UpdateGeoControls();
        };

        // Режим для слабого ПК: меньше одновременных скачиваний (число можно поправить вручную)
        chkEconomy.CheckedChanged += (_, _) => numThreads.Value = chkEconomy.Checked ? 4 : 16;

        btnRecode.Click += (_, _) =>
        {
            using var dlg = new RecoderForm(recode);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            recode = dlg.Result;
            UpdateRecodeLabel();
        };

        btnStart.Click += async (_, _) => await StartAsync();

        btnStop.Click += (_, _) =>
        {
            if (cts == null) return;
            btnStop.Enabled = false;
            log.AppendText("Останавливаю... (подождите несколько секунд)" + Environment.NewLine);
            cts.Cancel();
        };
    }

    private void UpdateRecodeLabel() => lblRecode.Text = recode.Describe();

    // Добавляет путь или ссылку в очередь (без повторов)
    private bool AddSource(string source)
    {
        source = source.Trim();
        if (source.Length == 0) return false;
        foreach (var existing in lstSources.Items)
            if (string.Equals((string)existing, source, StringComparison.OrdinalIgnoreCase)) return false;
        lstSources.Items.Add(source);
        return true;
    }

    // Добавляет то, что набрано в поле. Пустое поле — не ошибка; неверный путь — сообщение и false.
    private bool AddFromTextBox()
    {
        string text = txtJson.Text.Trim().Trim('"');   // кавычки остаются при вставке пути через «Копировать как путь»
        if (text.Length == 0) return true;

        if (!Converter.IsUrl(text) && !File.Exists(text))
        {
            MessageBox.Show("Укажите существующий JSON-файл или ссылку (http:// или https://).", "Ошибка");
            return false;
        }

        AddSource(text);
        txtJson.Clear();
        ApplySuggestion();
        return true;
    }

    // Имя для папки результата по файлу или ссылке:
    //  - файл   -> имя файла без расширения;
    //  - ссылка -> последняя часть ссылки (или адрес сайта).
    private static string SourceName(string source)
    {
        source = source.Trim();
        string name = "result";

        if (Converter.IsUrl(source))
        {
            try
            {
                var u = new Uri(source);
                var seg = u.Segments.Select(s => s.Trim('/')).LastOrDefault(s => s.Length > 0);
                name = !string.IsNullOrEmpty(seg) ? Path.GetFileNameWithoutExtension(seg) : u.Host;
            }
            catch { }
        }
        else
        {
            name = Path.GetFileNameWithoutExtension(source);
        }

        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        return string.IsNullOrWhiteSpace(name) ? "result" : name;
    }

    // Один JSON: папка рядом с файлом (для ссылки — в «Документах») с таким же именем.
    private static string? SuggestOutSingle(string source)
    {
        if (Converter.IsUrl(source))
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), SourceName(source));

        if (File.Exists(source))
            return Path.Combine(Path.GetDirectoryName(source)!, SourceName(source));

        return null;
    }

    // Несколько JSON: общая папка «Результат», внутри — подпапка на каждый файл.
    private string? SuggestOut()
    {
        var sources = lstSources.Items.Cast<string>().ToList();
        if (sources.Count == 0) return null;
        if (sources.Count == 1) return SuggestOutSingle(sources[0]);

        var firstFile = sources.FirstOrDefault(s => !Converter.IsUrl(s) && File.Exists(s));
        string dir = firstFile != null
            ? Path.GetDirectoryName(firstFile)!
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return Path.Combine(dir, "Результат");
    }

    // Папка результата всегда следует за списком. Если нужна своя папка — выберите её после добавления файлов.
    private void ApplySuggestion()
    {
        var s = SuggestOut();
        if (s != null) txtOut.Text = s;
    }

    // Фильтр «минимум в N% магазинов» не имеет смысла, когда выгрузка делается по одному магазину
    private void UpdateGeoControls()
    {
        bool perShop = cmbRegion.SelectedIndex == (int)RegionMode.PerShop;
        chkGeo.Enabled = numGeo.Enabled = !isBusy && !perShop;
    }

    private void SetBusy(bool busy)
    {
        isBusy = busy;
        btnStart.Enabled = btnJson.Enabled = btnAddLink.Enabled = btnRemove.Enabled = btnClear.Enabled = btnOut.Enabled = !busy;
        txtJson.Enabled = lstSources.Enabled = btnRecode.Enabled = chkEconomy.Enabled = numThreads.Enabled = !busy;
        btnStop.Enabled = busy;
        numSpread.Enabled = cmbPrice.Enabled = numPartMb.Enabled = chkZip.Enabled = numLimit.Enabled = chkMerge.Enabled = cmbRegion.Enabled = !busy;
        UpdateGeoControls();
    }

    private async Task StartAsync()
    {
        if (!AddFromTextBox()) return;   // то, что набрано в поле, добавляем в очередь

        if (lstSources.Items.Count == 0)
        {
            MessageBox.Show("Добавьте хотя бы один JSON-файл или ссылку.", "Ошибка");
            return;
        }

        var sources = lstSources.Items.Cast<string>().ToList();
        foreach (var s in sources)
        {
            if (!Converter.IsUrl(s) && !File.Exists(s))
            {
                MessageBox.Show($"Файл не найден:\n{s}", "Ошибка");
                return;
            }
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
        bool shopsOn = chkGeo.Checked;
        decimal shopsPct = numGeo.Value;
        var regionMode = (RegionMode)cmbRegion.SelectedIndex;   // порядок пунктов совпадает с enum RegionMode
        bool economy = chkEconomy.Checked;
        int threads = (int)numThreads.Value;
        var recodeOptions = recode;
        string baseDir = txtOut.Text.Trim();

        // Один JSON — результат прямо в выбранной папке; несколько — в подпапках по имени файла
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var items = new List<ConvertOptions>();
        foreach (var src in sources)
        {
            string dir = baseDir;
            if (sources.Count > 1)
            {
                string name = SourceName(src);
                string unique = name;
                int k = 2;
                while (!usedNames.Add(unique)) unique = $"{name} ({k++})";
                dir = Path.Combine(baseDir, unique);
            }

            items.Add(new ConvertOptions(
                src, dir, spread, mode, partMb, zip, limit, merge, shopsOn, shopsPct, regionMode,
                recodeOptions, threads, economy));
        }

        cts = new CancellationTokenSource();
        var token = cts.Token;

        // Режим для слабого ПК: программа работает с пониженным приоритетом, чтобы компьютер не «ложился»
        var process = System.Diagnostics.Process.GetCurrentProcess();
        var oldPriority = process.PriorityClass;
        if (economy)
        {
            try { process.PriorityClass = System.Diagnostics.ProcessPriorityClass.BelowNormal; } catch { }
        }

        // Progress<T> создан в UI-потоке, поэтому обработчики выполняются тоже в нём
        var logProgress = new Progress<string>(s => log.AppendText(s + Environment.NewLine));
        var percentProgress = new Progress<int>(p => progress.Value = Math.Clamp(p, 0, 100));

        try
        {
            string result = await Task.Run(() =>
                Converter.RunBatchAsync(items, logProgress, percentProgress, token));

            progress.Value = 100;
            MessageBox.Show($"Готово!\n\n{result}\n\nПапка: {baseDir}", "JSON → Excel");
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
            if (economy)
            {
                try { process.PriorityClass = oldPriority; } catch { }
            }
            cts?.Dispose();
            cts = null;
            SetBusy(false);
        }
    }
}

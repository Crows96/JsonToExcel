namespace JsonToExcel;

/// <summary>Элемент очереди: файл или ссылка, необязательное название (магазина / папки) и регион.</summary>
public sealed record SourceItem(string Source, string? Label = null, string? Region = null)
{
    public override string ToString() =>
        (Label == null ? Source : $"{Label}   ←   {Source}") + (Region == null ? "" : $"   [{Region}]");
}

public class MainForm : Form
{
    private readonly TextBox txtJson = new()
    {
        Dock = DockStyle.Fill, PlaceholderText = "Путь к файлу (JSON или YML/XML) или ссылка https://... (Enter — добавить)"
    };
    private readonly Button btnJson = new() { Text = "Файлы...", AutoSize = true };
    private readonly Button btnAddLink = new() { Text = "Добавить", AutoSize = true };
    private readonly Button btnList = new() { Text = "Список из txt...", AutoSize = true };
    private readonly ListBox lstSources = new()
    {
        Dock = DockStyle.Fill, SelectionMode = SelectionMode.MultiExtended, HorizontalScrollbar = true,
        IntegralHeight = false, AllowDrop = true
    };
    private readonly Button btnRemove = new() { Text = "Убрать", AutoSize = true };
    private readonly Button btnClear = new() { Text = "Очистить", AutoSize = true };
    private readonly CheckBox chkCombine = new()
    {
        Text = "Объединить все файлы очереди в один результат (YML/XML и JSON)",
        AutoSize = true, Margin = new Padding(0, 4, 0, 4)
    };
    private readonly TextBox txtOut = new() { Dock = DockStyle.Fill };
    private readonly Button btnOut = new() { Text = "Обзор...", AutoSize = true };
    private readonly NumericUpDown numSpread = new()
    {
        DecimalPlaces = 1, Minimum = 0, Maximum = 100, Increment = 0.5m, Value = 0, Width = 70,
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
    private readonly CheckBox chkSubject = new()
    {
        Text = "Регионом считать субъект РФ (область, край, республика), а не город", AutoSize = true,
        Margin = new Padding(3, 4, 3, 2)
    };
    private readonly CheckBox chkPromo = new()
    {
        Text = "Только акционные позиции (есть price_old — старая цена)", AutoSize = true,
        Margin = new Padding(20, 4, 0, 0)
    };
    private readonly CheckBox chkGeo = new()
    {
        Text = "Фильтр по магазинам (target_shops): товар минимум в", Checked = true, AutoSize = true, Margin = new Padding(0, 4, 4, 0)
    };
    private readonly NumericUpDown numGeo = new()
    {
        Minimum = 1, Maximum = 100, Value = 95, Width = 60, Margin = new Padding(0, 3, 4, 0)
    };
    private readonly ComboBox cmbRegion = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList, Width = 480, Margin = new Padding(0, 3, 0, 0)
    };
    private readonly Label lblRegionHint = new()
    {
        AutoSize = true, MaximumSize = new Size(700, 0), ForeColor = SystemColors.GrayText, Margin = new Padding(3, 6, 3, 2)
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

    private readonly ComboBox cmbPreset = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
    private readonly Button btnPresetSave = new() { Text = "Сохранить как...", AutoSize = true };
    private readonly Button btnPresetDelete = new() { Text = "Удалить", AutoSize = true };
    private readonly Button btnPresetReset = new() { Text = "По умолчанию", AutoSize = true };
    private readonly Button btnAdvanced = new() { Text = "▸ Дополнительные настройки", AutoSize = true };
    private readonly FlowLayoutPanel advancedPanel = new()
    {
        AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Visible = false,
        Margin = new Padding(0, 0, 0, 4)
    };
    private readonly Label lblSummary = new() { AutoSize = true, MaximumSize = new Size(740, 0) };
    private SettingsFile store = new();
    private bool settingsExisted;
    private bool applying;   // true, пока настройки применяются программно: обработчики «сами подстраивают» значения только от действий человека
    private string? listDir;   // папка последнего загруженного txt-списка (для папки результата)
    private bool isBusy;

    public MainForm()
    {
        Text = "JSON / YML → Excel + картинки";
        Width = 800;
        Height = 900;
        MinimumSize = new Size(800, 700);
        StartPosition = FormStartPosition.CenterScreen;

        cmbPrice.Items.AddRange(new object[] { "Минимальная", "Максимальная", "Средняя", "Медиана" });
        cmbPrice.SelectedIndex = 1;   // по умолчанию — максимальная цена

        // Порядок пунктов совпадает с enum RegionMode
        cmbRegion.Items.AddRange(new object[]
        {
            "Не разбивать — один Excel на все магазины",
            "Только товары, которые есть во всех регионах — один Excel",
            "По регионам — отдельный Excel (папка) на каждый регион / город",
            "По магазинам — отдельный Excel (папка) на каждый магазин",
            "По регионам, но один общий Excel — фильтр считается в каждом регионе отдельно"
        });
        cmbRegion.SelectedIndex = 0;
        UpdateRegionHint();

        UpdateRecodeLabel();

        var sourceButtons = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        sourceButtons.Controls.Add(btnJson);
        sourceButtons.Controls.Add(btnList);
        sourceButtons.Controls.Add(btnAddLink);

        var listButtons = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.TopDown };
        listButtons.Controls.Add(btnRemove);
        listButtons.Controls.Add(btnClear);

        store = SettingsStore.Load(out settingsExisted);

        var tip = new ToolTip { AutoPopDelay = 30000, InitialDelay = 300, ReshowDelay = 100 };

        // Подпись с всплывающей подсказкой
        Label Caption(string text, string hint)
        {
            var label = new Label { Text = text, AutoSize = true, Margin = new Padding(0, 6, 6, 0) };
            tip.SetToolTip(label, hint);
            return label;
        }

        // Строка из нескольких элементов
        FlowLayoutPanel RowOf(params Control[] controls)
        {
            var panel = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 0, 2) };
            panel.Controls.AddRange(controls);
            return panel;
        }

        // Рамка с заголовком и строками внутри
        GroupBox Group(string title, params Control[] rows)
        {
            var box = new GroupBox
            {
                Text = title, Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 4, 8, 6), Margin = new Padding(0, 4, 0, 4)
            };
            var inner = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false
            };
            inner.Controls.AddRange(rows);
            box.Controls.Add(inner);
            return box;
        }

        // Понятные подписи (поля создаются выше, здесь только тексты)
        chkPromo.Text = "Только акционные позиции (есть старая цена: price_old в JSON, oldprice в YML)";
        chkPromo.Margin = new Padding(0, 4, 0, 0);
        chkGeo.Text = "Брать только товары, которые есть минимум в";

        // ----- Подсказки -----
        tip.SetToolTip(txtJson, "Путь к файлу (JSON или YML/XML) или ссылка. Enter — добавить в очередь.");
        tip.SetToolTip(btnJson, "Выбрать один или несколько файлов. Файлы можно и просто перетащить в очередь.");
        tip.SetToolTip(btnList, "Загрузить список файлов и ссылок из txt. Одна строка — один файл: ссылка;название магазина;регион. " +
                                "Название и регион необязательны.");
        tip.SetToolTip(btnAddLink, "Добавить в очередь то, что набрано в поле слева.");
        tip.SetToolTip(lstSources, "Очередь обработки. Файлы обрабатываются по очереди.");
        tip.SetToolTip(chkSubject, "Если включено, регион определяется по справочнику: город из адреса магазина заменяется его областью, краем или республикой " +
                                   "(Химки и Мытищи — «Московская область», Зеленоград — «Москва»), а названия из target_regions приводятся к официальным. " +
                                   "Если выключено, регионом остаётся город из адреса. Нужен встроенный справочник ru_geo.tsv.");
        tip.SetToolTip(chkCombine, "Все файлы очереди считаются одним набором. В YML каждый файл — один магазин. В JSON магазины берутся " +
                                   "из target_shops внутри файлов; если один и тот же магазин встретился в нескольких файлах, он " +
                                   "учитывается один раз, а один и тот же товар — тоже один раз. В результат попадут товары, которые " +
                                   "повторяются в нужном проценте магазинов. Получится один Excel (или несколько частей, если картинок много).");
        tip.SetToolTip(txtOut, "Папка, куда будут сохранены результаты. По умолчанию — рядом с файлом.");
        tip.SetToolTip(chkPromo, "Берём только акционные товары — те, у которых указана старая цена.");
        tip.SetToolTip(chkGeo, "Оставляем только товары, которые есть минимум в указанном проценте магазинов. Например, 80 — товар должен " +
                               "быть минимум в 8 из 10 магазинов. В JSON магазины берутся из target_shops, при объединении — это файлы очереди.");
        tip.SetToolTip(numGeo, "Минимальный процент магазинов, в которых должен быть товар.");
        tip.SetToolTip(numSpread, "Если один товар в разных магазинах стоит по-разному, берём его только если разница между самой " +
                                  "высокой и самой низкой ценой не больше этого значения. 0 — цена должна совпадать. " +
                                  "Пример: 5 — цены 100 и 104 подходят, 100 и 110 — нет.");
        tip.SetToolTip(cmbPrice, "Если у товара несколько цен (в разных магазинах), какую из них записать в Excel.");
        tip.SetToolTip(cmbRegion, "Как разделить результат: один общий Excel, Excel только с товарами, которые есть во всех регионах, " +
                                  "или отдельная папка на каждый регион или магазин. Названия берутся из target_regions и target_shops. Если в файле нет target_regions, регион определяется как город из адреса магазина.");
        tip.SetToolTip(btnRecode, "Уменьшить и пережать картинки при скачивании: файлы становятся в разы легче, частей получается меньше.");
        tip.SetToolTip(numPartMb, "Результат делится на части (part_001, part_002...), чтобы архивы с картинками не были слишком большими. " +
                                  "Здесь — максимальный размер картинок в одной части.");
        tip.SetToolTip(chkZip, "Картинки каждой части упаковываются в один ZIP-архив (так удобнее загружать). " +
                               "Без галочки они лежат в папке images.");
        tip.SetToolTip(numLimit, "Для пробного запуска: взять только первые N товаров. 0 — без ограничения.");
        tip.SetToolTip(chkMerge, "Если один и тот же товар встречается под разными кодами, но с одинаковым названием и картинкой, " +
                                 "считать его одним товаром.");
        tip.SetToolTip(chkEconomy, "Меньше нагрузки на компьютер: 4 потока скачивания, одно сжатие картинок за раз, программа работает " +
                                   "с пониженным приоритетом. Работает медленнее.");
        tip.SetToolTip(numThreads, "Сколько картинок скачивается одновременно. Меньше — медленнее, но легче для слабого компьютера и интернета.");
        tip.SetToolTip(cmbPreset, "Готовый набор настроек. Выберите набор, чтобы применить его.");
        tip.SetToolTip(btnPresetSave, "Сохранить текущие настройки под именем, например «Магнит».");
        tip.SetToolTip(btnPresetDelete, "Удалить выбранный набор.");
        tip.SetToolTip(btnPresetReset, "Вернуть рекомендуемые настройки.");

        // ----- Блоки настроек -----
        var presetRow = RowOf(Caption("Набор настроек:", "Сохраните удачные настройки под именем и выбирайте их одним щелчком."),
                              cmbPreset, btnPresetSave, btnPresetDelete, btnPresetReset);

        var productsGroup = Group("Какие товары брать",
            RowOf(chkPromo),
            RowOf(chkGeo, numGeo, Caption("% магазинов", "Минимальный процент магазинов, в которых должен быть товар.")),
            RowOf(Caption("Допустимая разница цен между магазинами, %:", "0 — цена должна совпадать. Пример: 5 — цены 100 и 104 подходят, 100 и 110 нет."),
                  numSpread),
            RowOf(Caption("Какую цену записать в Excel:", "Если у товара несколько цен, какую из них записать."), cmbPrice));

        // Отдельный заметный блок: для больших фидов по всей стране именно здесь выбирается «по регионам»
        var regionGroup = Group("Разбивка по регионам и магазинам",
            RowOf(Caption("Как разбить результат:", "Один общий Excel или отдельные папки по регионам и магазинам."), cmbRegion),
            lblRegionHint,
            chkSubject);

        var imagesGroup = Group("Картинки",
            RowOf(Caption("Сжатие:", "Уменьшение и пережатие картинок при скачивании."), lblRecode, btnRecode),
            RowOf(Caption("Размер одной части, МБ:", "Максимальный размер картинок в одной части результата."), numPartMb, chkZip));

        advancedPanel.Controls.Add(RowOf(Caption("Ограничить число товаров (0 — без ограничения):", "Для пробного запуска: взять только первые N товаров."), numLimit));
        advancedPanel.Controls.Add(RowOf(chkMerge));
        advancedPanel.Controls.Add(RowOf(chkEconomy, Caption("Потоков скачивания:", "Сколько картинок скачивается одновременно."), numThreads));

        var summaryGroup = Group("Что будет сделано", lblSummary);

        var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        buttons.Controls.Add(btnStart, 0, 0);
        buttons.Controls.Add(btnStop, 1, 0);

        // ----- Раскладка окна -----
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(10) };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        int row = 0;

        // Строка на всю ширину окна
        void AddWide(Control control, RowStyle style)
        {
            layout.RowStyles.Add(style);
            layout.Controls.Add(control, 0, row);
            layout.SetColumnSpan(control, 3);
            row++;
        }

        // Строка из трёх ячеек: подпись, поле, кнопки
        void AddTriple(Control caption, Control field, Control buttonsCell, RowStyle style)
        {
            layout.RowStyles.Add(style);
            layout.Controls.Add(caption, 0, row);
            layout.Controls.Add(field, 1, row);
            layout.Controls.Add(buttonsCell, 2, row);
            row++;
        }

        AddTriple(new Label { Text = "Добавить файл:", AutoSize = true, Anchor = AnchorStyles.Left },
                  txtJson, sourceButtons, new RowStyle(SizeType.AutoSize));
        AddTriple(new Label { Text = "Очередь (по очереди):", AutoSize = true, Anchor = AnchorStyles.Top | AnchorStyles.Left, Margin = new Padding(3, 6, 3, 0) },
                  lstSources, listButtons, new RowStyle(SizeType.Absolute, 100));
        AddWide(chkCombine, new RowStyle(SizeType.AutoSize));
        AddTriple(new Label { Text = "Папка результата:", AutoSize = true, Anchor = AnchorStyles.Left },
                  txtOut, btnOut, new RowStyle(SizeType.AutoSize));

        AddWide(presetRow, new RowStyle(SizeType.AutoSize));
        AddWide(productsGroup, new RowStyle(SizeType.AutoSize));
        AddWide(regionGroup, new RowStyle(SizeType.AutoSize));
        AddWide(imagesGroup, new RowStyle(SizeType.AutoSize));
        AddWide(btnAdvanced, new RowStyle(SizeType.AutoSize));
        AddWide(advancedPanel, new RowStyle(SizeType.AutoSize));
        AddWide(summaryGroup, new RowStyle(SizeType.AutoSize));
        AddWide(buttons, new RowStyle(SizeType.Absolute, 44));
        AddWide(progress, new RowStyle(SizeType.Absolute, 28));
        AddWide(log, new RowStyle(SizeType.Percent, 100));
        layout.RowCount = row;

        Controls.Add(layout);

        // ----- Дополнительные настройки: показать / скрыть -----
        btnAdvanced.Click += (_, _) =>
        {
            advancedPanel.Visible = !advancedPanel.Visible;
            btnAdvanced.Text = advancedPanel.Visible ? "▾ Дополнительные настройки" : "▸ Дополнительные настройки";
        };

        // ----- Наборы настроек -----
        cmbPreset.SelectedIndexChanged += (_, _) =>
        {
            if (applying || cmbPreset.SelectedItem is not string name) return;
            if (store.Presets.TryGetValue(name, out var preset)) ApplySettings(preset);
        };

        btnPresetSave.Click += (_, _) =>
        {
            string? name = AskText(this, "Сохранить набор настроек", "Название набора (например, «Галамарт»):", cmbPreset.Text);
            if (string.IsNullOrWhiteSpace(name)) return;
            if (store.Presets.ContainsKey(name) &&
                MessageBox.Show($"Набор «{name}» уже есть. Заменить?", "Набор настроек", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;
            store.Presets[name] = CaptureSettings();
            SettingsStore.Save(store);
            RefreshPresetList(name);
        };

        btnPresetDelete.Click += (_, _) =>
        {
            if (cmbPreset.SelectedItem is not string name) return;
            if (MessageBox.Show($"Удалить набор «{name}»?", "Набор настроек", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            store.Presets.Remove(name);
            SettingsStore.Save(store);
            RefreshPresetList();
        };

        btnPresetReset.Click += (_, _) =>
        {
            ApplySettings(new UiSettings());
            RefreshPresetList();
        };

        // ----- «Что будет сделано» обновляется при любом изменении настроек -----
        foreach (var box in new CheckBox[] { chkPromo, chkGeo, chkZip, chkEconomy, chkCombine, chkMerge })
            box.CheckedChanged += (_, _) => UpdateSummary();
        foreach (var number in new NumericUpDown[] { numGeo, numSpread, numPartMb, numLimit, numThreads })
            number.ValueChanged += (_, _) => UpdateSummary();
        foreach (var combo in new ComboBox[] { cmbPrice, cmbRegion })
            combo.SelectedIndexChanged += (_, _) => UpdateSummary();

        // Настройки запоминаются при закрытии окна
        FormClosed += (_, _) =>
        {
            store.Last = CaptureSettings();
            SettingsStore.Save(store);
        };

        // Возвращаем настройки прошлого запуска
        RefreshPresetList();
        if (settingsExisted) ApplySettings(store.Last);
        UpdateSummary();

        // ----- Список JSON -----
        btnJson.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "JSON и YML (*.json;*.xml;*.yml)|*.json;*.xml;*.yml|Все файлы|*.*", Multiselect = true };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            foreach (var f in dlg.FileNames) AddSource(f);
            ApplySuggestion();
        };

        btnAddLink.Click += (_, _) => AddFromTextBox();

        btnList.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "Список ссылок (*.txt;*.csv)|*.txt;*.csv|Все файлы|*.*" };
            if (dlg.ShowDialog() == DialogResult.OK) LoadList(dlg.FileName);
        };

        // Объединение всегда строится по охвату магазинов — включаем этот фильтр (его можно выключить вручную)
        chkCombine.CheckedChanged += (_, _) => { if (!applying && chkCombine.Checked) chkGeo.Checked = true; };

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

        btnClear.Click += (_, _) => { lstSources.Items.Clear(); UpdateSummary(); };

        // Файлы можно просто перетащить в список
        lstSources.DragEnter += (_, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy;
        };
        lstSources.DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] files) return;
            foreach (var f in files)
            {
                // txt и csv — это список ссылок, остальное — сами файлы
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext is ".txt" or ".csv") LoadList(f); else AddSource(f);
            }
            ApplySuggestion();
        };

        btnOut.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK) txtOut.Text = dlg.SelectedPath;
        };

        // Охват по магазинам и регионам считается по объединённым товарам, поэтому включаем объединение автоматически
        chkGeo.CheckedChanged += (_, _) => { if (!applying && chkGeo.Checked) chkMerge.Checked = true; };
        cmbRegion.SelectedIndexChanged += (_, _) =>
        {
            // Для режимов по регионам охват считается по объединённым товарам; для магазинов объединение не обязательно
            if (!applying && cmbRegion.SelectedIndex is (int)RegionMode.AllRegions or (int)RegionMode.PerRegion or (int)RegionMode.RegionsMerged)
                chkMerge.Checked = true;
            UpdateGeoControls();
            UpdateRegionHint();
        };

        // Режим для слабого ПК: меньше одновременных скачиваний (число можно поправить вручную)
        chkEconomy.CheckedChanged += (_, _) => { if (!applying) numThreads.Value = chkEconomy.Checked ? 4 : 16; };

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

    private void UpdateRecodeLabel()
    {
        lblRecode.Text = recode.Describe();
        UpdateSummary();
    }

    // Склонение по числу: 1 поток, 2 потока, 5 потоков
    private static string Plural(int n, string one, string few, string many)
    {
        int mod100 = n % 100, mod10 = n % 10;
        if (mod100 is >= 11 and <= 14) return many;
        if (mod10 == 1) return one;
        if (mod10 is >= 2 and <= 4) return few;
        return many;
    }

    // Есть ли в очереди JSON (локальные файлы определяются по содержимому, ссылки — по расширению)
    private bool QueueHasJson() =>
        lstSources.Items.Cast<SourceItem>().Any(s =>
        {
            if (Converter.IsUrl(s.Source)) return s.Source.Contains(".json", StringComparison.OrdinalIgnoreCase);
            try { return File.Exists(s.Source) && Converter.DetectFormat(s.Source) == SourceFormat.Json; }
            catch { return false; }
        });

    // Простыми словами описывает, что произойдёт при нажатии «Старт»
    private void UpdateSummary()
    {
        int n = lstSources.Items.Count;
        bool combine = chkCombine.Checked;
        int geoMode = cmbRegion.SelectedIndex;
        bool perShop = geoMode == (int)RegionMode.PerShop;
        bool shopFilter = chkGeo.Checked && !perShop;
        decimal spread = numSpread.Value;
        var lines = new List<string>();

        if (n == 0) lines.Add("Добавьте в очередь файлы или ссылки (JSON, YML/XML).");
        else if (combine && n < 2) lines.Add("⚠ Для объединения добавьте в очередь минимум два файла.");
        else if (combine) lines.Add($"Файлов в очереди: {n}. Из них будет собран один результат. " +
                                    (QueueHasJson()
                                        ? "Магазины берутся из самих JSON-файлов (target_shops); повторяющиеся магазины и товары учитываются один раз."
                                        : "Каждый файл — отдельный магазин."));
        else if (n == 1) lines.Add("Будет обработан один файл, результат — в выбранной папке.");
        else lines.Add($"Файлов в очереди: {n}. Каждый обрабатывается отдельно и получает свою папку.");

        lines.Add(chkPromo.Checked
            ? "Берутся только акционные позиции (есть старая цена)."
            : "Берутся все позиции, не только акционные.");

        if (geoMode == (int)RegionMode.AllRegions)
            lines.Add("Остаются только товары, которые есть во всех регионах (target_regions; если их нет — город из адреса магазина).");
        else if (geoMode == (int)RegionMode.PerRegion)
            lines.Add("Для каждого региона (target_regions; если их нет — город из адреса магазина) создаётся своя папка с Excel.");
        else if (geoMode == (int)RegionMode.RegionsMerged)
            lines.Add("Фильтр считается в каждом регионе отдельно (регион — target_regions или город из адреса), результат — один общий Excel.");
        else if (perShop)
            lines.Add("Для каждого магазина или адреса (target_shops) создаётся своя папка с Excel.");

        if (shopFilter)
        {
            decimal pct = numGeo.Value;
            if (combine && n >= 2 && QueueHasJson())
                lines.Add($"Товар должен быть минимум в {pct:0.##}% магазинов из всех файлов. Сколько их всего, программа посчитает " +
                          "при проверке файлов перед чтением и покажет в журнале.");
            else if (combine && n >= 2)
            {
                int need = (int)Math.Ceiling(n * pct / 100m);
                lines.Add($"Товар должен быть минимум в {pct:0.##}% магазинов — то есть в {need} из {n}.");
            }
            else
                lines.Add($"Товар должен быть минимум в {pct:0.##}% магазинов (по target_shops).");
        }

        // Про цены имеет смысл говорить, только когда у товара может быть несколько цен
        if (combine || shopFilter || geoMode > 0)
        {
            lines.Add(spread == 0
                ? "Цена товара во всех магазинах должна совпадать."
                : $"Цены в разных магазинах могут отличаться не более чем на {spread:0.##}%.");
            if (spread > 0)
            {
                string[] names = { "минимальная", "максимальная", "средняя", "медианная" };
                lines.Add($"В Excel записывается {names[Math.Clamp(cmbPrice.SelectedIndex, 0, 3)]} цена.");
            }
        }

        if (numLimit.Value > 0) lines.Add($"Пробный запуск: не больше {numLimit.Value:0} товаров.");

        lines.Add(recode.Enabled
            ? $"Картинки сжимаются: {recode.Describe()}."
            : "Картинки сохраняются без сжатия.");
        lines.Add($"Результат делится на части: до {numPartMb.Value:0} МБ картинок в каждой" +
                  (chkZip.Checked ? ", картинки упаковываются в ZIP." : ", картинки лежат в папке images."));

        if (chkEconomy.Checked)
        {
            int t = (int)numThreads.Value;
            lines.Add($"Режим для слабого ПК: {t} {Plural(t, "поток", "потока", "потоков")} скачивания, пониженная нагрузка на компьютер.");
        }

        lblSummary.Text = string.Join(Environment.NewLine, lines.Select(l => l.StartsWith('⚠') ? l : "• " + l));
    }

    // Текущие настройки окна -> объект для сохранения
    private UiSettings CaptureSettings() => new()
    {
        Spread = numSpread.Value,
        PriceMode = Math.Max(0, cmbPrice.SelectedIndex),
        PartMb = (int)numPartMb.Value,
        Zip = chkZip.Checked,
        Limit = (int)numLimit.Value,
        Merge = chkMerge.Checked,
        PromoOnly = chkPromo.Checked,
        ShopFilter = chkGeo.Checked,
        ShopPercent = numGeo.Value,
        GeoMode = Math.Max(0, cmbRegion.SelectedIndex),
        RegionBySubject = chkSubject.Checked,
        Combine = chkCombine.Checked,
        Economy = chkEconomy.Checked,
        Threads = (int)numThreads.Value,
        RecodeEnabled = recode.Enabled,
        RecodeFormat = (int)recode.Format,
        RecodeMaxSide = recode.MaxSide,
        RecodeQuality = recode.Quality
    };

    // Сохранённые настройки -> окно (значения приводятся к допустимым пределам)
    private void ApplySettings(UiSettings s)
    {
        applying = true;
        try
        {
            numSpread.Value = Math.Clamp(s.Spread, numSpread.Minimum, numSpread.Maximum);
            cmbPrice.SelectedIndex = Math.Clamp(s.PriceMode, 0, cmbPrice.Items.Count - 1);
            numPartMb.Value = Math.Clamp(s.PartMb, 1, 100000);
            chkZip.Checked = s.Zip;
            numLimit.Value = Math.Clamp(s.Limit, 0, 1000000);
            chkMerge.Checked = s.Merge;
            chkPromo.Checked = s.PromoOnly;
            chkGeo.Checked = s.ShopFilter;
            numGeo.Value = Math.Clamp(s.ShopPercent, numGeo.Minimum, numGeo.Maximum);
            cmbRegion.SelectedIndex = Math.Clamp(s.GeoMode, 0, cmbRegion.Items.Count - 1);
            chkSubject.Checked = s.RegionBySubject;
            chkCombine.Checked = s.Combine;
            chkEconomy.Checked = s.Economy;
            numThreads.Value = Math.Clamp(s.Threads, 1, 64);
            recode = new RecodeOptions(s.RecodeEnabled, (RecodeFormat)Math.Clamp(s.RecodeFormat, 0, 2),
                                       Math.Clamp(s.RecodeMaxSide, 0, 10000), Math.Clamp(s.RecodeQuality, 1, 100));
        }
        finally
        {
            applying = false;
        }

        UpdateRecodeLabel();
        UpdateGeoControls();
        UpdateSummary();
    }

    // Обновляет список наборов (без применения выбранного)
    private void RefreshPresetList(string? select = null)
    {
        applying = true;
        try
        {
            cmbPreset.Items.Clear();
            foreach (var name in store.Presets.Keys.OrderBy(k => k, StringComparer.CurrentCultureIgnoreCase))
                cmbPreset.Items.Add(name);
            if (select != null) cmbPreset.SelectedItem = select;
        }
        finally
        {
            applying = false;
        }
    }

    // Небольшое окно с одной строкой ввода
    private static string? AskText(IWin32Window owner, string title, string prompt, string initial = "")
    {
        using var form = new Form
        {
            Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false, ClientSize = new Size(420, 110)
        };
        var label = new Label { Text = prompt, AutoSize = true, Location = new Point(12, 12) };
        var box = new TextBox { Text = initial, Location = new Point(12, 36), Width = 396 };
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(232, 72), Width = 85 };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, Location = new Point(323, 72), Width = 85 };
        form.Controls.AddRange(new Control[] { label, box, ok, cancel });
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        return form.ShowDialog(owner) == DialogResult.OK ? box.Text.Trim() : null;
    }

    // Добавляет элемент в очередь (без повторов по пути или ссылке)
    private bool AddItem(SourceItem item)
    {
        string source = item.Source.Trim();
        if (source.Length == 0) return false;
        foreach (var existing in lstSources.Items)
            if (string.Equals(((SourceItem)existing).Source, source, StringComparison.OrdinalIgnoreCase)) return false;
        lstSources.Items.Add(item with { Source = source });
        return true;
    }

    private bool AddSource(string source) => AddItem(new SourceItem(source));

    // Загружает список из txt/csv. Одна строка — один файл или ссылка:
    //   ссылка
    //   ссылка;название магазина (станет названием папки или магазина)
    //   ссылка;название магазина;регион
    // Разделитель — «;» или табуляция. Пустые строки и строки, начинающиеся с «#», пропускаются.
    private void LoadList(string path)
    {
        string text;
        try
        {
            byte[] bytes = File.ReadAllBytes(path);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            try { text = new System.Text.UTF8Encoding(false, true).GetString(bytes); }
            catch (System.Text.DecoderFallbackException) { text = System.Text.Encoding.GetEncoding(1251).GetString(bytes); }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Не удалось прочитать список: " + ex.Message, "Ошибка");
            return;
        }
        if (text.Length > 0 && text[0] == '\uFEFF') text = text[1..];

        string listFolder = Path.GetDirectoryName(Path.GetFullPath(path)) ?? "";
        int added = 0, duplicates = 0;
        var bad = new List<string>();

        lstSources.BeginUpdate();
        try
        {
            foreach (var raw in text.Split('\n'))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith('#')) continue;

                var parts = line.Split(';', '\t');
                string source = parts[0].Trim().Trim('"');
                string? label = parts.Length > 1 && parts[1].Trim().Length > 0 ? parts[1].Trim() : null;
                string? region = parts.Length > 2 && parts[2].Trim().Length > 0 ? parts[2].Trim() : null;

                // Относительные пути считаем от папки, где лежит сам список
                if (!Converter.IsUrl(source) && !Path.IsPathRooted(source))
                    source = Path.Combine(listFolder, source);

                if (!Converter.IsUrl(source) && !File.Exists(source)) { bad.Add(line); continue; }

                if (AddItem(new SourceItem(source, label, region))) added++; else duplicates++;
            }
        }
        finally
        {
            lstSources.EndUpdate();
        }

        listDir = listFolder;
        ApplySuggestion();

        string summary = $"Список {Path.GetFileName(path)}: добавлено {added}" +
                         (duplicates > 0 ? $", повторов пропущено {duplicates}" : "") +
                         (bad.Count > 0 ? $", не найдено {bad.Count}" : "");
        log.AppendText(summary + Environment.NewLine);

        if (bad.Count > 0)
            MessageBox.Show(summary + "\n\nНе найдены (первые 8):\n" + string.Join("\n", bad.Take(8)), "Список ссылок");
    }

    // Добавляет то, что набрано в поле. Пустое поле — не ошибка; неверный путь — сообщение и false.
    private bool AddFromTextBox()
    {
        string text = txtJson.Text.Trim().Trim('"');   // кавычки остаются при вставке пути через «Копировать как путь»
        if (text.Length == 0) return true;

        if (!Converter.IsUrl(text) && !File.Exists(text))
        {
            MessageBox.Show("Укажите существующий файл (JSON или YML/XML) или ссылку (http:// или https://).", "Ошибка");
            return false;
        }

        AddSource(text);
        txtJson.Clear();
        ApplySuggestion();
        return true;
    }

    // Имя для папки результата по файлу или ссылке:
    //  - YML-файл -> название магазина из файла (shop/name), если оно там есть;
    //  - файл     -> имя файла без расширения;
    //  - ссылка   -> последняя часть ссылки (или адрес сайта).
    private static string SourceName(string source, bool useShopName = true)
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
            name = (useShopName ? Converter.TryReadShopName(source) : null) ?? Path.GetFileNameWithoutExtension(source);
        }

        return SafeName(name);
    }

    // Делает из названия допустимое имя папки
    private static string SafeName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        name = name.Trim().TrimEnd('.');
        if (name.Length > 80) name = name[..80].TrimEnd();
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
        var sources = lstSources.Items.Cast<SourceItem>().Select(i => i.Source).ToList();
        if (sources.Count == 0) return null;
        if (sources.Count == 1) return SuggestOutSingle(sources[0]);

        var firstFile = sources.FirstOrDefault(s => !Converter.IsUrl(s) && File.Exists(s));
        string dir = firstFile != null
            ? Path.GetDirectoryName(firstFile)!
            : listDir ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return Path.Combine(dir, "Результат");
    }

    // Папка результата всегда следует за списком. Если нужна своя папка — выберите её после добавления файлов.
    private void ApplySuggestion()
    {
        UpdateSummary();
        var s = SuggestOut();
        if (s != null) txtOut.Text = s;
    }

    // Фильтр «минимум в N% магазинов» не имеет смысла, когда выгрузка делается по одному магазину
    private void UpdateGeoControls()
    {
        bool perShop = cmbRegion.SelectedIndex == (int)RegionMode.PerShop;
        chkGeo.Enabled = numGeo.Enabled = !isBusy && !perShop;
        chkSubject.Enabled = !isBusy && cmbRegion.SelectedIndex is (int)RegionMode.AllRegions or (int)RegionMode.PerRegion or (int)RegionMode.RegionsMerged;
    }

    // Короткое пояснение под выбором разбивки: что будет сделано в выбранном режиме
    private void UpdateRegionHint()
    {
        lblRegionHint.Text = (RegionMode)Math.Max(0, cmbRegion.SelectedIndex) switch
        {
            RegionMode.AllRegions =>
                "Остаются только товары, которые есть в каждом регионе. Регион — это target_regions, а если их в файле нет — город из адреса магазина (target_shops).",
            RegionMode.PerRegion =>
                "Для каждого региона создаётся своя папка с Excel. Регион — это target_regions, а если их в файле нет — город из адреса магазина (target_shops). " +
                "Процент магазинов и разброс цен считаются внутри региона. Рекомендуется для больших фидов по всей стране.",
            RegionMode.PerShop =>
                "Для каждого магазина (адреса) создаётся своя папка с Excel. Фильтр по проценту магазинов не применяется.",
            RegionMode.RegionsMerged =>
                "Результат — один Excel. Процент магазинов и разброс цен проверяются в каждом регионе отдельно (регион — это target_regions, " +
                "а если их нет — город из адреса магазина). В файл попадает товар, который прошёл проверку в каждом регионе, где он есть. " +
                "Цена берётся по всем магазинам по правилу «Какую цену записать».",
            _ =>
                "Все магазины считаются вместе. Для фида по всей стране процент магазинов считается от всех магазинов сразу, " +
                "поэтому товаров может остаться мало — для таких фидов выберите «По регионам»."
        };
    }

    private void SetBusy(bool busy)
    {
        isBusy = busy;
        btnStart.Enabled = btnJson.Enabled = btnList.Enabled = btnAddLink.Enabled = btnRemove.Enabled = btnClear.Enabled = btnOut.Enabled = !busy;
        chkCombine.Enabled = !busy;
        cmbPreset.Enabled = btnPresetSave.Enabled = btnPresetDelete.Enabled = btnPresetReset.Enabled = !busy;
        txtJson.Enabled = lstSources.Enabled = btnRecode.Enabled = chkEconomy.Enabled = numThreads.Enabled = !busy;
        btnStop.Enabled = busy;
        numSpread.Enabled = cmbPrice.Enabled = numPartMb.Enabled = chkZip.Enabled = numLimit.Enabled = chkMerge.Enabled = chkPromo.Enabled = cmbRegion.Enabled = !busy;
        UpdateGeoControls();
    }

    private async Task StartAsync()
    {
        if (!AddFromTextBox()) return;   // то, что набрано в поле, добавляем в очередь

        if (lstSources.Items.Count == 0)
        {
            MessageBox.Show("Добавьте хотя бы один файл (JSON или YML/XML) или ссылку.", "Ошибка");
            return;
        }

        var sources = lstSources.Items.Cast<SourceItem>().ToList();
        foreach (var s in sources)
        {
            if (!Converter.IsUrl(s.Source) && !File.Exists(s.Source))
            {
                MessageBox.Show($"Файл не найден:\n{s.Source}", "Ошибка");
                return;
            }
        }

        if (chkCombine.Checked)
        {
            if (sources.Count < 2)
            {
                MessageBox.Show("Для объединения нужно минимум два файла в очереди.", "Ошибка");
                return;
            }
        }

        if (string.IsNullOrWhiteSpace(txtOut.Text)) ApplySuggestion();
        if (string.IsNullOrWhiteSpace(txtOut.Text))
        {
            MessageBox.Show("Укажите папку для результата.", "Ошибка");
            return;
        }

        store.Last = CaptureSettings();
        SettingsStore.Save(store);

        SetBusy(true);
        log.Clear();
        progress.Value = 0;

        decimal spread = numSpread.Value;
        var mode = (PriceMode)cmbPrice.SelectedIndex;   // порядок пунктов совпадает с enum PriceMode
        int partMb = (int)numPartMb.Value;
        bool zip = chkZip.Checked;
        int limit = (int)numLimit.Value;
        bool merge = chkMerge.Checked;
        bool promoOnly = chkPromo.Checked;
        bool shopsOn = chkGeo.Checked;
        decimal shopsPct = numGeo.Value;
        var regionMode = (RegionMode)cmbRegion.SelectedIndex;   // порядок пунктов совпадает с enum RegionMode
        bool economy = chkEconomy.Checked;
        int threads = (int)numThreads.Value;
        var recodeOptions = recode;
        string baseDir = txtOut.Text.Trim();

        var items = new List<ConvertOptions>();

        if (chkCombine.Checked)
        {
            // Объединение: один результат в выбранной папке, каждый файл очереди — отдельный магазин
            var feeds = sources
                .Select(s => new FeedSource(s.Source, s.Label ?? SourceName(s.Source, useShopName: false), s.Region))
                .ToList();
            items.Add(new ConvertOptions(
                sources[0].Source, baseDir, spread, mode, partMb, zip, limit, merge, shopsOn, shopsPct, regionMode,
                recodeOptions, threads, economy, promoOnly, feeds, chkSubject.Checked));
        }
        else
        {
            // Один файл — результат прямо в выбранной папке; несколько — в подпапках по названию
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var src in sources)
            {
                string dir = baseDir;
                if (sources.Count > 1)
                {
                    string name = src.Label != null ? SafeName(src.Label) : SourceName(src.Source);
                    string unique = name;
                    int k = 2;
                    while (!usedNames.Add(unique)) unique = $"{name} ({k++})";
                    dir = Path.Combine(baseDir, unique);
                }

                items.Add(new ConvertOptions(
                    src.Source, dir, spread, mode, partMb, zip, limit, merge, shopsOn, shopsPct, regionMode,
                    recodeOptions, threads, economy, promoOnly, null, chkSubject.Checked));
            }
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
            // Полоса убирается (сбрасывается), и строка в журнале говорит, что программа остановилась
            progress.Value = 0;
            log.AppendText("ОШИБКА: " + ex.Message + Environment.NewLine);
            log.AppendText("■ Программа остановлена, больше ничего не выполняется. Измените настройки и запустите снова." + Environment.NewLine);
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

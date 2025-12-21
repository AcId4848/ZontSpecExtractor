using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZontSpecExtractor.Properties;
using static ZontSpecExtractor.Form1;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.VisualBasic;
using System.Diagnostics.Metrics;


namespace ZontSpecExtractor
{

    // =========================================================================
    // 1. КОНФИГУРАЦИЯ И НАСТРОЙКИ (Сохранение состояния)
    // =========================================================================

    // Вне класса VisioConfiguration (но в том же namespace)
    public record PredefinedMasterConfig
    {
        // === ЭТО ВСТАВИТЬ В НАЧАЛО КЛАССА (ВНЕ МЕТОДОВ) ===
        public string MasterName { get; set; } = "";
        public int Quantity { get; set; } = 1;
        public string CoordinatesXY { get; set; } = ""; // Пример: "10,20"
        public string Anchor { get; set; } = "Center";
        public double Width { get; set; } = 1.0;
        public double Height { get; set; } = 1.0;
    }

    // Новый класс для хранения данных мастера
    public record MasterData(string MasterName, double WidthMM, double HeightMM)
    {
        public override string ToString() => $"{MasterName} ({WidthMM:F1}x{HeightMM:F1} мм)";
    }


    public class SequentialDrawingConfig
    {
        public bool Enabled { get; set; } = true;
        public string StartCoordinatesXY { get; set; } = "10, 250"; // Start X, Start Y
        public double MaxLineWidthMM { get; set; } = 180.0;
        public double HorizontalStepMM { get; set; } = 2.0;
        public double VerticalStepMM { get; set; } = 10.0; // Междустрочный интервал
        public string Anchor { get; set; } = "TopLeft";
    }

    public class VisioConfiguration
    {
        public List<string> StencilFilePaths { get; set; } = new List<string>();

        public List<string> AvailableMasters { get; set; } = new List<string>();

        // Основной список правил маппинга
        public List<SearchRule> SearchRules { get; set; } = new List<SearchRule>();

        // Список заранее определенных фигур
        public List<PredefinedMasterConfig> PredefinedMasterConfigs { get; set; } = new List<PredefinedMasterConfig>();

        public SequentialDrawingConfig SequentialDrawing { get; set; } = new SequentialDrawingConfig();




        public string PageSize { get; set; } = "A4"; // Например: "A4", "A3"
        public string PageOrientation { get; set; } = "Portrait"; // Или "Landscape"

        public VisioConfiguration(bool isScheme)
        {
            SearchRules = new List<SearchRule>();

            if (isScheme)
            {
                SearchRules.Add(new SearchRule
                {
                    ExcelValue = "H2000+proV2",
                    VisioMasterName = "H2000+PRO ZONT Контроллер",
                    SearchColumn = "C", // Ищем только в колонке C
                    UseCondition = true
                });

                SearchRules.Add(new SearchRule
                {
                    ExcelValue = "H1500+pro",
                    VisioMasterName = "H1500+PRO ZONT Контроллер",
                    SearchColumn = "C",
                    UseCondition = true
                });
            }
            else
            {
                // Пример для другого режима
                SearchRules.Add(new SearchRule
                {
                    ExcelValue = "H2000+proV2",
                    VisioMasterName = "Маркировка H2000 (пример)",
                    SearchColumn = "C",
                    UseCondition = true
                });
            }
        }

        public VisioConfiguration() : this(true) { }
    }

    public enum DataSourceType
    {
        MainFile, // Основная таблица (где ищем "1")
        AuxFile   // Вторая таблица (спецификация/доп данные)
    }

    public class SearchRule
    {
        // === ОСНОВНЫЕ ДАННЫЕ ДЛЯ ПОИСКА ===

        // То, что ищем в Excel (может быть "Слово1; Слово2")
        public string ExcelValue { get; set; } = "";

        // Колонка, в которой ищем совпадение (например "V" или "A"). 
        // Если пусто — ищем во всех колонках (или по логике UseCondition)
        public string SearchColumn { get; set; } = "";

        // То, что вставляем в Visio (может быть "Master1; Master2")
        public string VisioMasterName { get; set; } = "";

        // === НАСТРОЙКИ РАЗМЕЩЕНИЯ ===

        public string CoordinatesXY { get; set; } = ""; // Например "10,20"
        public string Anchor { get; set; } = "Center";  // Точка привязки (TopLeft, Center...)
        public bool LimitQuantity { get; set; } = false; // Если true, то берем только 1 раз

        // === ДОПОЛНИТЕЛЬНЫЕ УСЛОВИЯ (Legacy / Фильтры) ===
        // Если у вас есть логика "Проверять колонку L на значение 1", оставляем это:
        public bool UseCondition { get; set; } = false;
        public string ConditionColumn { get; set; } = "L";
        public string ConditionValue { get; set; } = "1";

        public string ExcludedRows { get; set; } = ""; // Например: "10, 37, 229-"

        public string TargetColumn { get; set; } = "Position";

        public bool SearchByValue { get; set; } = false; // Галочка "Искать по значению"

        // Свойство для совместимости, если где-то используется old-style "Term"
        public string Term
        {
            get => ExcelValue;
            set => ExcelValue = value;
        }

        public DataSourceType ResultSource { get; set; } = DataSourceType.MainFile;
    }



    public class SearchConfiguration
    {
        public List<string> TargetSheetNames { get; set; } = new List<string> { "1.ТЗ на объект ZONT" };
        public List<string> VisioSourceSheetName { get; set; } = new List<string> { "1.ТЗ на объект ZONT для VISIO" };
        // Заменяем простой список строк на список правил
        public List<SearchRule> Rules { get; set; } = new List<SearchRule>();

        public string AuxFilePath { get; set; } = "";
    }

    public static class AppSettings
    {
        private const string SettingsFilePath = "app_settings.json";

        public static VisioConfiguration SchemeConfig { get; private set; } = new VisioConfiguration();
        public static VisioConfiguration LabelingConfig { get; private set; } = new VisioConfiguration();
        // 2. НОВОЕ: Настройки для Шкафа
        public static VisioConfiguration CabinetConfig { get; private set; } = new VisioConfiguration();

        public static SearchConfiguration SearchConfig { get; private set; } = new SearchConfiguration();
        public static string LastLoadedFilePath { get; set; } = "";

        public static void Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonSerializer.Deserialize<Dictionary<string, object>>(json);

                    if (settings != null)
                    {
                        if (settings.TryGetValue("Scheme", out var s)) SchemeConfig = JsonSerializer.Deserialize<VisioConfiguration>(s.ToString()) ?? new VisioConfiguration();
                        if (settings.TryGetValue("Labeling", out var l)) LabelingConfig = JsonSerializer.Deserialize<VisioConfiguration>(l.ToString()) ?? new VisioConfiguration();
                        // Загрузка Шкафа
                        if (settings.TryGetValue("Cabinet", out var c)) CabinetConfig = JsonSerializer.Deserialize<VisioConfiguration>(c.ToString()) ?? new VisioConfiguration();

                        if (settings.TryGetValue("Search", out var sr)) SearchConfig = JsonSerializer.Deserialize<SearchConfiguration>(sr.ToString()) ?? new SearchConfiguration();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine($"Config Load Error: {ex.Message}"); }
        }

        public static void Save()
        {
            try
            {
                var settings = new Dictionary<string, object>
                {
                    { "Scheme", SchemeConfig },
                    { "Labeling", LabelingConfig },
                    { "Cabinet", CabinetConfig }, // Сохранение Шкафа
                    { "Search", SearchConfig }
                };
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFilePath, json);
            }
            catch { }
        }

        // Хелпер для получения конфига по типу
        public static VisioConfiguration GetConfig(string type)
        {
            switch (type)
            {
                case "Scheme": return SchemeConfig;
                case "Labeling": return LabelingConfig;
                case "Cabinet": return CabinetConfig;
                default: return SchemeConfig;
            }
        }
    }



    // =========================================================================
    // 2. ФОРМА ОБЩИХ НАСТРОЕК (GeneralSettingsForm)
    // =========================================================================


    public class GeneralSettingsForm : Form
    {
        // Поля для старых вкладок
        private RichTextBox _rtxtSchemePaths, _rtxtLabelingPaths, _rtxtCabinetPaths;
        private RichTextBox _rtxtSchemeMap, _rtxtLabelingMap;
        private RichTextBox _rtxtCabinetMap;
        // Predefined текстбоксы оставим только для совместимости, но редактировать будем через GUI
        private RichTextBox _rtxtSchemePredefined, _rtxtLabelingPredefined, _rtxtCabinetPredefined;

        private RichTextBox _rtxtSheetNames;
        private DataGridView _dgvSearchRules;

        public GeneralSettingsForm()
        {
            this.Text = "Общие настройки ZONT Extractor";
            this.Size = new Size(1200, 850);
            this.StartPosition = FormStartPosition.CenterParent;
            SetupUI();
        }

        private void SetupUI()
        {
            var mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), RowCount = 2, RowStyles = { new RowStyle(SizeType.Percent, 100), new RowStyle(SizeType.Absolute, 50) } };
            var tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9F) };

            // 1. Вкладка Поиска
            var searchPage = new TabPage("1. Настройки поиска (Excel)") { Padding = new Padding(10) };
            searchPage.Controls.Add(CreateSearchConfigPanel());
            tabControl.TabPages.Add(searchPage);

            // 2. Вкладка Visio (Трафареты и Маппинг)
            var visioPage = new TabPage("2. Настройки Visio (Трафареты)") { Padding = new Padding(10) };
            visioPage.Controls.Add(CreateVisioConfigPanel());
            tabControl.TabPages.Add(visioPage);

            // 3. НОВАЯ ВКЛАДКА: Настройки схем (Визуал)
            var schemaPage = new TabPage("3. Настройки схем (Визуал)") { Padding = new Padding(10) };
            schemaPage.Controls.Add(CreateSchemaVisualPanel());
            tabControl.TabPages.Add(schemaPage);

            mainLayout.Controls.Add(tabControl, 0, 0);

            // Кнопки
            var btnSave = new Button { Text = "Сохранить всё", Width = 150, Height = 35, DialogResult = DialogResult.OK, BackColor = Color.LightGreen };
            btnSave.Click += (s, e) =>
            {
                // === ВАЖНОЕ ИСПРАВЛЕНИЕ ===
                // 1. Принудительно забираем данные из таблицы обратно в конфиг
                if (_dgvSearchRules.DataSource is System.ComponentModel.BindingList<SearchRule> list)
                {
                    AppSettings.SearchConfig.Rules = list.ToList();
                }

                // 2. Сохраняем настройки
                AppSettings.Save();

                // Закрываем форму
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            var btnCancel = new Button { Text = "Отмена", Width = 100, Height = 35, DialogResult = DialogResult.Cancel };

            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            footer.Controls.Add(btnCancel);
            footer.Controls.Add(btnSave);
            mainLayout.Controls.Add(footer, 0, 1);
            this.Controls.Add(mainLayout);
        }

        // --- ПАНЕЛЬ 3: ВИЗУАЛЬНЫЕ НАСТРОЙКИ СХЕМ ---
        private Control CreateSchemaVisualPanel()
        {
            var tabs = new TabControl { Dock = DockStyle.Fill };

            // Создаем 3 вкладки для каждого типа листа
            tabs.TabPages.Add(CreateSingleSheetSettingsPage("МАРКИРОВКА", AppSettings.LabelingConfig));
            tabs.TabPages.Add(CreateSingleSheetSettingsPage("СХЕМА", AppSettings.SchemeConfig));
            tabs.TabPages.Add(CreateSingleSheetSettingsPage("ШКАФ", AppSettings.CabinetConfig));

            return tabs;
        }

        // Это должен быть метод внутри класса GeneralSettingsForm
        private TabPage CreateSingleSheetSettingsPage(string title, VisioConfiguration cfg)
        {
            var page = new TabPage(title);
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 600 };

            // --- ЛЕВАЯ ЧАСТЬ: НАСТРОЙКИ ---
            var leftPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };

            // 1. Верх: Размер и Ориентация
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 60 };
            var cbSize = new ComboBox { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
            cbSize.Items.AddRange(new object[] { "A4", "A3" });
            cbSize.SelectedItem = cfg.PageSize ?? "A4";

            var btnOrient = new Button { Text = cfg.PageOrientation == "Landscape" ? "Альбомная" : "Книжная", Width = 100, BackColor = Color.LightYellow };

            topPanel.Controls.Add(new Label { Text = "Размер:", Location = new Point(5, 10), AutoSize = true });
            topPanel.Controls.Add(cbSize); cbSize.Location = new Point(60, 8);
            topPanel.Controls.Add(new Label { Text = "Ориентация:", Location = new Point(160, 10), AutoSize = true });
            topPanel.Controls.Add(btnOrient); btnOrient.Location = new Point(240, 5);
            leftPanel.Controls.Add(topPanel);

            // 2. Группа: Найденные фигуры (Sequential)
            var grpSeq = new GroupBox { Text = "Настройки размещения (Змейка/Список)", Dock = DockStyle.Top, Height = 140 }; // Высоту можно подстроить
            var flowSeq = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoScroll = true };

            var txtStart = new TextBox { Text = cfg.SequentialDrawing.StartCoordinatesXY, Width = 80 };
            var numMaxW = new NumericUpDown { Minimum = 0, Maximum = 1000, Value = (decimal)cfg.SequentialDrawing.MaxLineWidthMM, Width = 60 };
            var numVGap = new NumericUpDown { Minimum = 0, Maximum = 100, Value = (decimal)cfg.SequentialDrawing.VerticalStepMM, Width = 60 };
            var numHGap = new NumericUpDown { Minimum = 0, Maximum = 100, Value = (decimal)cfg.SequentialDrawing.HorizontalStepMM, Width = 60 };

            // Галочка включения
            var chkEn = new CheckBox { Text = "Включить", Checked = cfg.SequentialDrawing.Enabled, AutoSize = true };

            // --- НОВОЕ: ГЛОБАЛЬНЫЙ ЯКОРЬ ---
            var lblAnchor = new Label { Text = "Якорь:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) };
            var cbAnchor = new ComboBox { Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            // Добавляем основные варианты (как вы просили: 3 вида, но можно и больше)
            cbAnchor.Items.AddRange(new object[] { "TopLeft", "Center", "BottomLeft" });

            // Устанавливаем текущее значение (если пусто - Center)
            string currentAnchor = cfg.SequentialDrawing.Anchor;
            if (string.IsNullOrEmpty(currentAnchor)) currentAnchor = "Center";
            if (!cbAnchor.Items.Contains(currentAnchor)) cbAnchor.Items.Add(currentAnchor);
            cbAnchor.SelectedItem = currentAnchor;

            // Добавляем контролы в FlowLayout
            AddControl(flowSeq, "Старт X,Y:", txtStart);
            AddControl(flowSeq, "Макс. ширина:", numMaxW);
            AddControl(flowSeq, "Отступ снизу:", numVGap);
            AddControl(flowSeq, "Отступ сбоку:", numHGap);

            // Добавляем чекбокс и сразу за ним настройку якоря
            flowSeq.Controls.Add(chkEn);
            flowSeq.Controls.Add(lblAnchor); // Подпись
            flowSeq.Controls.Add(cbAnchor);  // Выпадающий список

            grpSeq.Controls.Add(flowSeq);

            // Добавляем настройки в левую панель
            leftPanel.Controls.Add(grpSeq);
            leftPanel.Controls.Add(topPanel);

            // 3. Таблица ФИКСИРОВАННЫХ фигур (Predefined)
            // Просто занимает всё оставшееся место (без SplitContainer)
            var pnlFixed = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 0) };
            var lblFixed = new Label { Text = "Фиксированные фигуры (Шапки, Рамки):", Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 9, FontStyle.Bold) };

            var dgvFixed = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                ColumnHeadersHeight = 30
            };
            dgvFixed.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Имя мастера", DataPropertyName = "MasterName", Width = 150 });
            dgvFixed.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "X (мм)", DataPropertyName = "X", Width = 50 });
            dgvFixed.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Y (мм)", DataPropertyName = "Y", Width = 50 });

            var colAnchorFixed = new DataGridViewComboBoxColumn { HeaderText = "Anchor", DataPropertyName = "Anchor", Width = 80 };
            colAnchorFixed.Items.AddRange("Center", "TopLeft", "BottomLeft");
            dgvFixed.Columns.Add(colAnchorFixed);

            var bindingListFixed = new System.ComponentModel.BindingList<PredefinedViewModel>(
                cfg.PredefinedMasterConfigs.Select(p => new PredefinedViewModel(p)).ToList()
            );
            dgvFixed.DataSource = bindingListFixed;

            pnlFixed.Controls.Add(dgvFixed);
            pnlFixed.Controls.Add(lblFixed);

            // Важно: pnlFixed добавляем первым (или с BringToFront), так как у него Dock.Fill
            // Но так как topPanel и grpSeq имеют Dock.Top, правильнее добавлять pnlFixed в leftPanel ПОСЛЕ них,
            // и WinForms заполнит оставшееся пространство.
            leftPanel.Controls.Add(pnlFixed);
            pnlFixed.BringToFront(); // На всякий случай, чтобы корректно занял место в центре

            split.Panel1.Controls.Add(leftPanel);

            // --- ПРАВАЯ ЧАСТЬ: ПРЕВЬЮ ---
            var previewPanel = new PagePreviewPanel
            {
                Dock = DockStyle.Fill,
                PageSize = cfg.PageSize,
                Orientation = cfg.PageOrientation,
                FixedShapes = cfg.PredefinedMasterConfigs,
                SeqConfig = cfg.SequentialDrawing
            };
            split.Panel2.Controls.Add(previewPanel);

            // --- СОБЫТИЯ ОБНОВЛЕНИЯ ---
            Action updatePreview = () =>
            {
                cfg.PageSize = cbSize.Text;
                cfg.SequentialDrawing.StartCoordinatesXY = txtStart.Text;
                cfg.SequentialDrawing.MaxLineWidthMM = (double)numMaxW.Value;
                cfg.SequentialDrawing.VerticalStepMM = (double)numVGap.Value;
                cfg.SequentialDrawing.HorizontalStepMM = (double)numHGap.Value;
                cfg.SequentialDrawing.Enabled = chkEn.Checked;

                // Сохраняем выбранный якорь
                if (cbAnchor.SelectedItem != null)
                    cfg.SequentialDrawing.Anchor = cbAnchor.SelectedItem.ToString();

                // Сохраняем фиксированные
                cfg.PredefinedMasterConfigs = bindingListFixed.Select(vm => vm.ToConfig()).ToList();

                previewPanel.PageSize = cfg.PageSize;
                previewPanel.Orientation = cfg.PageOrientation;
                previewPanel.FixedShapes = cfg.PredefinedMasterConfigs;
                previewPanel.SeqConfig = cfg.SequentialDrawing;

                previewPanel.Invalidate();
            };

            // Привязка событий
            cbSize.SelectedIndexChanged += (s, e) => updatePreview();
            btnOrient.Click += (s, e) => {
                cfg.PageOrientation = (cfg.PageOrientation == "Portrait") ? "Landscape" : "Portrait";
                btnOrient.Text = cfg.PageOrientation == "Landscape" ? "Альбомная" : "Книжная";
                previewPanel.Orientation = cfg.PageOrientation;
                updatePreview();
            };

            txtStart.TextChanged += (s, e) => updatePreview();
            numMaxW.ValueChanged += (s, e) => updatePreview();
            numVGap.ValueChanged += (s, e) => updatePreview();
            numHGap.ValueChanged += (s, e) => updatePreview();
            chkEn.CheckedChanged += (s, e) => updatePreview();

            // При изменении якоря обновляем конфиг
            cbAnchor.SelectedIndexChanged += (s, e) => updatePreview();

            dgvFixed.CellEndEdit += (s, e) => updatePreview();
            dgvFixed.RowsRemoved += (s, e) => updatePreview();
            dgvFixed.UserAddedRow += (s, e) => updatePreview();

            page.Controls.Add(split);
            return page;
        }

        // Вспомогательный метод для добавления контролов
        private void AddControl(FlowLayoutPanel panel, string label, Control ctrl)
        {
            var p = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Margin = new Padding(5) };
            p.Controls.Add(new Label { Text = label, AutoSize = true });
            p.Controls.Add(ctrl);
            panel.Controls.Add(p);
        }
        // Это должен быть метод внутри класса GeneralSettingsForm
        private void UpdatePreviewFromGrid(PagePreviewPanel panel, System.ComponentModel.BindingList<PredefinedViewModel> list, VisioConfiguration cfg)
        {
            // Конвертируем ViewModel обратно в Config для превью и основного конфига
            var configs = list.Select(vm => vm.ToConfig()).ToList();

            // Обновляем основной конфиг (для сохранения)
            cfg.PredefinedMasterConfigs = configs;

            // Обновляем превью
            panel.FixedShapes = configs;
            panel.Invalidate();
        }

        // ViewModel для удобного редактирования в гриде (разделяем X и Y)
        public class PredefinedViewModel
        {
            public string MasterName { get; set; }
            public string X { get; set; } // Ввод в MM
            public string Y { get; set; } // Ввод в MM
            public int Quantity { get; set; } = 1;
            public string Anchor { get; set; } = "Center";

            // --- ИСПРАВЛЕНИЕ: Добавляем хранение размеров ---
            public double Width { get; set; } = 30.0;
            public double Height { get; set; } = 15.0;
            // ------------------------------------------------

            public PredefinedViewModel() { }
            public PredefinedViewModel(PredefinedMasterConfig cfg)
            {
                MasterName = cfg.MasterName;
                Quantity = cfg.Quantity;

                // Сохраняем реальные размеры из конфига
                Width = cfg.Width;
                Height = cfg.Height;

                string cleanedAnchor = cfg.Anchor?.Trim();
                if (string.IsNullOrWhiteSpace(cleanedAnchor)) Anchor = "Center";
                else if (cleanedAnchor.Equals("Center", StringComparison.OrdinalIgnoreCase)) Anchor = "Center";
                else if (cleanedAnchor.Equals("TopLeft", StringComparison.OrdinalIgnoreCase)) Anchor = "TopLeft";
                else if (cleanedAnchor.Equals("BottomLeft", StringComparison.OrdinalIgnoreCase)) Anchor = "BottomLeft";
                else Anchor = "Center";

                var parts = cfg.CoordinatesXY?.Split(',');
                if (parts != null && parts.Length >= 2)
                {
                    X = parts[0].Trim();
                    Y = parts[1].Trim();
                }
                else { X = "0"; Y = "0"; }
            }

            public PredefinedMasterConfig ToConfig()
            {
                return new PredefinedMasterConfig
                {
                    MasterName = MasterName,
                    Quantity = Quantity,
                    CoordinatesXY = $"{X},{Y}",
                    Anchor = Anchor,
                    // --- ИСПРАВЛЕНИЕ: Возвращаем сохраненные размеры обратно в конфиг ---
                    Width = Width,
                    Height = Height
                    // --------------------------------------------------------------------
                };
            }
        }

        // --- КОНЕЦ НОВОЙ ПАНЕЛИ ---


        private Panel CreateSearchConfigPanel()
        {
            var panel = new Panel { Dock = DockStyle.Fill };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 5, // <-- БЫЛО 4, СТАЛО 5 (добавилась строка для файла)
                RowStyles = {
            new RowStyle(SizeType.Absolute, 30), // Заголовок листов
            new RowStyle(SizeType.Absolute, 60), // Поле листов
            new RowStyle(SizeType.Absolute, 40), // <-- НОВАЯ СТРОКА: Доп. таблица
            new RowStyle(SizeType.Absolute, 30), // Заголовок правил
            new RowStyle(SizeType.Percent, 100)  // Таблица правил
        }
            };

            // 1. Листы
            layout.Controls.Add(new Label { Text = "Целевые листы (откуда брать данные):", Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);
            _rtxtSheetNames = new RichTextBox { Dock = DockStyle.Fill, Text = string.Join(Environment.NewLine, AppSettings.SearchConfig.TargetSheetNames), ReadOnly = true, BackColor = SystemColors.ControlLight };
            layout.Controls.Add(_rtxtSheetNames, 0, 1);

            // === [НОВОЕ] 2. Окно выбора дополнительной таблицы ===
            var auxPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(0, 5, 0, 0) };
            var lblAux = new Label { Text = "Доп. таблица:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) };

            var txtAux = new TextBox { Width = 350, ReadOnly = true, BackColor = SystemColors.ControlLight };
            // Привязка к свойству AuxFilePath (убедитесь, что оно есть в AppSettings.SearchConfig)
            txtAux.DataBindings.Add("Text", AppSettings.SearchConfig, "AuxFilePath", false, DataSourceUpdateMode.OnPropertyChanged);

            var btnAux = new Button { Text = "...", Width = 30, Height = 23 };
            btnAux.Click += (s, e) =>
            {
                using var ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtAux.Text = ofd.FileName;
                    AppSettings.SearchConfig.AuxFilePath = ofd.FileName; // Явное обновление
                }
            };

            var btnClearAux = new Button { Text = "X", Width = 30, Height = 23, ForeColor = Color.Red };
            btnClearAux.Click += (s, e) => { txtAux.Text = ""; AppSettings.SearchConfig.AuxFilePath = ""; };

            auxPanel.Controls.AddRange(new Control[] { lblAux, txtAux, btnAux, btnClearAux });
            layout.Controls.Add(auxPanel, 0, 2);
            // =======================================================

            // 3. Заголовок Правила поиска
            layout.Controls.Add(new Label { Text = "Правила поиска и условий:", Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 3);

            _dgvSearchRules = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                ColumnHeadersHeight = 40,
                AllowUserToResizeColumns = true
            };

            // --- КОЛОНКИ ТАБЛИЦЫ (ВАШИ СТАРЫЕ) ---

            // 1. Искомое слово
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Искомое слово\n(Visio Key)",
                DataPropertyName = "ExcelValue",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 150
            });

            // 2. Где искать само слово
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Где искать\n(Col: C)",
                DataPropertyName = "SearchColumn",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            // 3. Галочка "Использовать условие"
            _dgvSearchRules.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "Есть\nусловие?",
                DataPropertyName = "UseCondition",
                Width = 60
            });

            // 4. Где искать условие
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Где условие\n(Col: L)",
                DataPropertyName = "ConditionColumn",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            // 5. Значение условия
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Значение\nусловия (1)",
                DataPropertyName = "ConditionValue",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            // 6. Фиксировать количество
            _dgvSearchRules.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "Всегда\n1 шт?",
                DataPropertyName = "LimitQuantity",
                Width = 60
            });

            // 7. Искать по значению (Галочка)
            _dgvSearchRules.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "Взять текст\nесли (1)?",
                DataPropertyName = "SearchByValue",
                Width = 80
            });

            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ExcludedRows",
                HeaderText = "Исключить строки (напр. 10, 229-)",
                Width = 120
            });

            var targetColParams = new DataGridViewComboBoxColumn
            {
                DataPropertyName = "TargetColumn",
                HeaderText = "Куда записать результат",
                Width = 150,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
            };
            // (Ваш список колонок TargetColumn)
            targetColParams.Items.AddRange(new object[] {
        "Position", "Col4", "Col5", "Col6", "Col7", "Col8",
        "Col9", "Col10", "Col11", "Col12", "Col13", "Col14",
        "Col15", "Col16", "Col17", "Col18"
    });
            _dgvSearchRules.Columns.Add(targetColParams);

            // === [НОВОЕ] 8. Откуда брать результат ===
            var sourceCol = new DataGridViewComboBoxColumn
            {
                HeaderText = "Откуда брать\nрезультат",
                DataPropertyName = "ResultSource",
                DataSource = Enum.GetValues(typeof(DataSourceType)),
                Width = 120,
                FlatStyle = FlatStyle.Flat
            };
            _dgvSearchRules.Columns.Add(sourceCol);
            // =========================================

            // Обработка ошибок DataError (чтобы не падало при выборе комбобокса)
            _dgvSearchRules.DataError += (s, e) => { e.ThrowException = false; };

            _dgvSearchRules.DataSource = new System.ComponentModel.BindingList<SearchRule>(AppSettings.SearchConfig.Rules ?? new List<SearchRule>());

            layout.Controls.Add(_dgvSearchRules, 0, 4); // <-- ИНДЕКС СТРОКИ СТАЛ 4

            panel.Controls.Add(layout);
            return panel;
        }

        private Panel CreateVisioConfigPanel()
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 2, RowStyles = { new RowStyle(SizeType.Absolute, 30), new RowStyle(SizeType.Percent, 100) }, ColumnStyles = { new ColumnStyle(SizeType.Percent, 33), new ColumnStyle(SizeType.Percent, 33), new ColumnStyle(SizeType.Percent, 33) } };
            layout.Controls.Add(new Label { Text = "Настройка трафаретов и маппинга имен", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);
            layout.SetColumnSpan(layout.GetControlFromPosition(0, 0), 3);

            _rtxtLabelingPaths = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = SystemColors.ControlLight };
            _rtxtLabelingMap = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };

            _rtxtSchemePaths = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = SystemColors.ControlLight };
            _rtxtSchemeMap = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };

            _rtxtCabinetPaths = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = SystemColors.ControlLight };
            _rtxtCabinetMap = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };

            // Инициализируем скрытые поля для совместимости
            _rtxtLabelingPredefined = new RichTextBox(); _rtxtSchemePredefined = new RichTextBox(); _rtxtCabinetPredefined = new RichTextBox();

            layout.Controls.Add(CreateSingleConfig("1. МАРКИРОВКА", AppSettings.LabelingConfig, _rtxtLabelingPaths, _rtxtLabelingMap), 0, 1);
            layout.Controls.Add(CreateSingleConfig("2. СХЕМА", AppSettings.SchemeConfig, _rtxtSchemePaths, _rtxtSchemeMap), 1, 1);
            layout.Controls.Add(CreateSingleConfig("3. ШКАФ", AppSettings.CabinetConfig, _rtxtCabinetPaths, _rtxtCabinetMap), 2, 1);

            return new Panel { Dock = DockStyle.Fill, Controls = { layout } };
        }

        private Panel CreateSingleConfig(string title, VisioConfiguration cfg, RichTextBox rPath, RichTextBox rMap)
        {
            var p = new Panel { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(2) };
            var l = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 7,
                RowStyles = {
            new RowStyle(SizeType.Absolute, 20),
            new RowStyle(SizeType.Absolute, 20),
            new RowStyle(SizeType.Absolute, 60),
            new RowStyle(SizeType.Absolute, 35),
            new RowStyle(SizeType.Percent, 40),
            new RowStyle(SizeType.Absolute, 20),
            new RowStyle(SizeType.Percent, 60)
        }
            };

            l.Controls.Add(new Label { Text = title, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);

            // 1. Трафареты
            l.Controls.Add(new Label { Text = "Трафареты:", Dock = DockStyle.Bottom }, 0, 1);
            rPath.Text = string.Join(Environment.NewLine, cfg.StencilFilePaths);
            l.Controls.Add(rPath, 0, 2);

            // 2. Кнопки (оставляем как было)
            var btnP = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            var btnAdd = new Button { Text = "+", Width = 30 };
            btnAdd.Click += (s, e) => {
                using (var ofd = new OpenFileDialog { Multiselect = true, Filter = "Visio|*.vssx;*.vsdx" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        cfg.StencilFilePaths.AddRange(ofd.FileNames.Where(f => !cfg.StencilFilePaths.Contains(f)));
                        rPath.Text = string.Join(Environment.NewLine, cfg.StencilFilePaths);
                    }
                }
            };
            var btnClear = new Button { Text = "X", Width = 30, ForeColor = Color.Red };
            btnClear.Click += (s, e) => { cfg.StencilFilePaths.Clear(); rPath.Text = ""; };

            // Кнопка Сканировать
            var btnScan = new Button { Text = "Сканировать", AutoSize = true, BackColor = Color.LightGreen };
            var rFoundMasters = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8), BackColor = Color.Beige };
            btnScan.Click += (s, e) => {
                ScanMasters(cfg);
                rFoundMasters.Text = string.Join(Environment.NewLine, cfg.AvailableMasters.OrderBy(m => m));
            };

            btnP.Controls.Add(btnAdd);
            btnP.Controls.Add(btnClear);
            btnP.Controls.Add(btnScan);
            l.Controls.Add(btnP, 0, 3);

            // 3. Найденные фигуры
            rFoundMasters.Text = string.Join(Environment.NewLine, cfg.AvailableMasters.OrderBy(m => m));
            var grpFound = new GroupBox { Text = "Найденные фигуры:", Dock = DockStyle.Fill };
            grpFound.Controls.Add(rFoundMasters);
            l.Controls.Add(grpFound, 0, 4);

            // =========================================================================
            // 4. ИЗМЕНЕНИЯ ЗДЕСЬ: КАРТА (Excel=Master)
            // =========================================================================
            l.Controls.Add(new Label { Text = "Карта (Excel=Master):", Dock = DockStyle.Bottom }, 0, 5);

            // Отображаем в формате "Excel=VisioMaster"
            rMap.Text = string.Join(Environment.NewLine, cfg.SearchRules.Select(r => $"{r.ExcelValue}={r.VisioMasterName}"));

            // Обработчик сохранения изменений из текста в конфиг
            rMap.Leave += (s, e) => {
                var newRules = new List<SearchRule>();
                foreach (var line in rMap.Lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split('=');
                    if (parts.Length >= 2)
                    {
                        newRules.Add(new SearchRule
                        {
                            // ВАЖНО: TrimStart() удаляет пробелы в начале, но оставляет в конце (например "Т ")
                            ExcelValue = parts[0].TrimStart(),
                            VisioMasterName = parts[1].Trim(),
                            SearchColumn = ""
                        });
                    }
                }
                cfg.SearchRules = newRules;
            };

            l.Controls.Add(rMap, 0, 6);
            p.Controls.Add(l);
            return p;
        }

        private void ScanMasters(VisioConfiguration cfg)
        {
            if (!cfg.StencilFilePaths.Any())
            {
                MessageBox.Show("Добавьте файлы трафаретов!", "Внимание"); return;
            }
            try
            {
                var masters = VisioMasterScanner.ScanStencils(cfg.StencilFilePaths);
                if (masters.Any())
                {
                    Clipboard.SetText(string.Join(Environment.NewLine, masters.OrderBy(m => m)));
                    MessageBox.Show($"Найдено {masters.Count} фигур. Список скопирован в буфер обмена.\nИспользуйте эти имена для настройки 'Фиксированных фигур' или 'Карты поиска'.", "Готово");
                }
                else MessageBox.Show("Фигуры не найдены.");
            }
            catch (Exception ex) { MessageBox.Show("Ошибка: " + ex.Message); }
        }

        private List<SearchRule> ParseSearchRules(RichTextBox rtb)
        {
            var rules = new List<SearchRule>();

            // Проверка на пустоту
            if (rtb == null || string.IsNullOrWhiteSpace(rtb.Text)) return rules;

            // Разбиваем текст на строки
            var lines = rtb.Text.Split(new[] { Environment.NewLine, "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var trimLine = line.Trim();
                // Пропускаем пустые строки и комментарии
                if (string.IsNullOrWhiteSpace(trimLine) || trimLine.StartsWith("//")) continue;

                // Разбиваем по знаку равенства
                var parts = trimLine.Split('=');

                // ВАРИАНТ 1: "Слово = Мастер" (2 части)
                if (parts.Length == 2)
                {
                    rules.Add(new SearchRule
                    {
                        ExcelValue = parts[0].Trim(),       // Что ищем в Excel
                        SearchColumn = "",                  // Колонка не указана (ищем везде или по дефолту)
                        VisioMasterName = parts[1].Trim()   // Имя фигуры в Visio
                    });
                }
                // ВАРИАНТ 2: "Слово = Колонка = Мастер" (3 части)
                else if (parts.Length == 3)
                {
                    rules.Add(new SearchRule
                    {
                        ExcelValue = parts[0].Trim(),       // Что ищем
                        SearchColumn = parts[1].Trim(),     // Конкретная колонка (например "B")
                        VisioMasterName = parts[2].Trim()   // Имя фигуры
                    });
                }
            }
            return rules;
        }

        private List<string> ParseList(RichTextBox rtb)
        {
            return rtb.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            try
            {
                // 1. Обновляем карты маппинга из текстовых полей
                AppSettings.SchemeConfig.SearchRules = ParseSearchRules(_rtxtSchemeMap);
                AppSettings.LabelingConfig.SearchRules = ParseSearchRules(_rtxtLabelingMap);
                AppSettings.CabinetConfig.SearchRules = ParseSearchRules(_rtxtCabinetMap);

                // (Фиксированные фигуры обновляются автоматически через BindingList в новой вкладке)

                // 2. Обновляем целевые листы
                AppSettings.SearchConfig.TargetSheetNames = ParseList(_rtxtSheetNames);

                // 3. Сохранение правил поиска
                var newRules = new List<SearchRule>();
                if (_dgvSearchRules.DataSource is System.ComponentModel.BindingList<SearchRule> list)
                {
                    newRules = list.Where(r => !string.IsNullOrWhiteSpace(r.Term)).ToList();
                }
                AppSettings.SearchConfig.Rules = newRules;

                AppSettings.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
            }
        }
    }




    // =========================================================================
    // 3. СКАНЕР МАСТЕРОВ VISIO (VisioMasterScanner)
    // =========================================================================

    public static class VisioMasterScanner
    {
        public static List<string> ScanStencils(List<string> filePaths)
        {
            var masters = new List<string>();
            Visio.Application? visioApp = null;
            bool createdNewApp = false;

            try
            {
                visioApp = new Visio.Application();
                createdNewApp = true;
                visioApp.Visible = false;

                foreach (var path in filePaths)
                {
                    if (!File.Exists(path)) continue;

                    Visio.Document? stencilDoc = null;
                    try
                    {
                        stencilDoc = visioApp.Documents.OpenEx(path, (short)Visio.VisOpenSaveArgs.visOpenDocked | (short)Visio.VisOpenSaveArgs.visOpenHidden);

                        foreach (Visio.Master master in stencilDoc.Masters)
                        {
                            masters.Add(master.NameU);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка при сканировании трафарета {Path.GetFileName(path)}: {ex.Message}");
                    }
                    finally
                    {
                        if (stencilDoc != null)
                        {
                            try { stencilDoc.Close(); } catch { }
                            Marshal.ReleaseComObject(stencilDoc);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Критическая ошибка при работе с Visio: {ex.Message}", "COM Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (visioApp != null)
                {
                    try
                    {
                        if (createdNewApp)
                        {
                            visioApp.Quit();
                        }
                    }
                    catch { }
                    Marshal.ReleaseComObject(visioApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return masters;
        }
    }


    // =========================================================================
    // 4. ОСНОВНАЯ ФОРМА (Form1)
    // =========================================================================
    // Вспомогательный класс для корректной работы с Visio Interop и освобождения COM-объектов
    public static class VisioScanner
    {
        // Метод для пакетного получения размеров. Решает проблему DOS handle.
        public static Dictionary<string, (double w, double h)> GetMastersDimensionsBatch(
            List<string> masterNames,
            List<string> stencilPaths)
        {
            var results = new Dictionary<string, (double, double)>(StringComparer.OrdinalIgnoreCase);
            var uniqueNames = new HashSet<string>(masterNames, StringComparer.OrdinalIgnoreCase);

            if (uniqueNames.Count == 0) return results;

            Visio.Application? visioApp = null;
            bool appCreated = false;

            try
            {
                // Пытаемся подключиться к открытому Visio или создаем новый
                try
                {
                    visioApp = (Visio.Application)Microsoft.VisualBasic.Interaction.GetObject("Visio.Application");
                }
                catch
                {
                    visioApp = new Visio.Application();
                    appCreated = true;
                }

                visioApp.Visible = false; // Работаем в фоне

                foreach (var path in stencilPaths)
                {
                    if (!File.Exists(path)) continue;

                    Visio.Document? stencil = null;
                    try
                    {
                        // Открываем трафарет
                        stencil = visioApp.Documents.OpenEx(path, (short)Visio.VisOpenSaveArgs.visOpenHidden);

                        // Проходимся по всем мастерам в трафарете
                        foreach (Visio.Master master in stencil.Masters)
                        {
                            // Если этот мастер есть в нашем списке запросов
                            if (uniqueNames.Contains(master.Name) || uniqueNames.Contains(master.NameU))
                            {
                                // Получаем размеры в миллиметрах
                                double w = master.Shapes[1].Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters];
                                double h = master.Shapes[1].Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                string key = master.Name; // Используем локальное имя
                                if (!results.ContainsKey(key))
                                {
                                    results[key] = (w, h);
                                }

                                // На всякий случай сохраняем и по NameU
                                if (!results.ContainsKey(master.NameU))
                                {
                                    results[master.NameU] = (w, h);
                                }
                            }
                        }
                    }
                    catch { /* Игнорируем битые файлы */ }
                    finally
                    {
                        if (stencil != null)
                        {
                            stencil.Close();
                            Marshal.ReleaseComObject(stencil);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сканирования Visio: " + ex.Message);
            }
            finally
            {
                if (appCreated && visioApp != null)
                {
                    visioApp.Quit();
                }
                if (visioApp != null) Marshal.ReleaseComObject(visioApp);
            }

            return results;
        }
    }

    public class RawExcelHit
    {
        public string SheetName { get; set; }
        public string FullItemName { get; set; }
        public string SearchTerm { get; set; }
        public bool ConditionMet { get; set; }
        public int Quantity { get; set; }
        public bool IsLimited { get; set; }

        public SearchRule FoundRule { get; set; }
        public string TargetMasterName { get; set; }

        public string TargetColumn { get; set; } = "Position";

        // === [ИНЪЕКЦИЯ 1] Поле для хранения приоритета (номера клеммы) ===
        public int SchematicOrder { get; set; } = 999;

        public int RowIndex { get; set; }
        public int SortIndex { get; set; }
    }



    public static class ExcelAnalysisHelper
    {
        /// <summary>
        /// Анализирует лист Excel, ищет строки, содержащие searchWord, и
        /// копирует их на новый лист, если условие в conditionColumnLetter равно 1.0.
        /// </summary>
        /// 

        public static void AnalyzeAndSaveSheet(
            string filePath,
            string sourceSheetName,
            string searchWord,
            string conditionColumnLetter)
        {
            // Убедитесь, что эта строка присутствует, если вы используете EPPlus (OfficeOpenXml)
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            FileInfo file = new FileInfo(filePath);

            // Используем FileAccess.ReadWrite, чтобы иметь возможность сохранить изменения
            using (var package = new ExcelPackage(file))
            {
                // 1. Получение исходного листа
                ExcelWorksheet sourceSheet = package.Workbook.Worksheets
                    .FirstOrDefault(s => s.Name.Equals(sourceSheetName, StringComparison.OrdinalIgnoreCase));

                if (sourceSheet == null)
                {
                    throw new Exception($"Лист '{sourceSheetName}' не найден. Проверьте правильность названия.");
                }

                // 2. Создание нового листа
                string newSheetName = "Анализ_" + searchWord;
                // Удаляем старый лист, если он существует, чтобы не было конфликта имен
                package.Workbook.Worksheets.Delete(newSheetName);
                ExcelWorksheet newSheet = package.Workbook.Worksheets.Add(newSheetName);
                int newRowIndex = 1;
                int copiedRowsCount = 0;

                // 3. Определение индекса колонки условия
                int conditionColumnIndex = 0; //ExcelColumnLetterToNumber(conditionColumnLetter);

                var dimension = sourceSheet.Dimension;
                if (dimension == null) return; // Лист пустой

                int startRow = dimension.Start.Row;
                int endRow = dimension.End.Row;
                int startCol = dimension.Start.Column;
                int endCol = dimension.End.Column;

                // Копирование заголовков (первой строки)
                for (int c = startCol; c <= endCol; c++)
                {
                    var sourceCell = sourceSheet.Cells[startRow, c];
                    var destCell = newSheet.Cells[newRowIndex, c];
                    destCell.Value = sourceCell.Value;
                    destCell.StyleID = sourceCell.StyleID;
                }
                if (startRow == 1) newRowIndex++; // Начинаем копирование данных со следующей строки

                // 4. Поиск и копирование строк
                for (int rCnt = startRow + 1; rCnt <= endRow; rCnt++)
                {
                    bool foundSearchWord = false;

                    // 4а. Ищем searchWord в текущей строке (итерация по ячейкам)
                    for (int c = startCol; c <= endCol; c++)
                    {
                        object cellValue = sourceSheet.Cells[rCnt, c].Value;
                        string cellText = cellValue?.ToString() ?? string.Empty;

                        // Поиск части строки без учета регистра
                        if (cellText.IndexOf(searchWord, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            foundSearchWord = true;
                            break;
                        }
                    }

                    if (foundSearchWord)
                    {
                        // 4б. Проверяем доп. условие: условие в колонке должно быть 1.0
                        bool conditionMet = false;
                        var conditionCell = sourceSheet.Cells[rCnt, conditionColumnIndex];
                        var conditionCellValue = conditionCell.Value;

                        if (conditionCellValue != null)
                        {
                            // Пытаемся получить числовое значение
                            if (double.TryParse(conditionCellValue.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedDouble))
                            {
                                // Сравнение с погрешностью
                                if (Math.Abs(parsedDouble - 1.0) < 0.001)
                                {
                                    conditionMet = true;
                                }
                            }
                        }

                        if (conditionMet)
                        {
                            // Копируем всю строку (значения и форматы)
                            for (int c = startCol; c <= endCol; c++)
                            {
                                var sourceCell = sourceSheet.Cells[rCnt, c];
                                var destCell = newSheet.Cells[newRowIndex, c];

                                // Копируем значение
                                destCell.Value = sourceCell.Value;

                                // Копируем форматирование
                                destCell.StyleID = sourceCell.StyleID;
                            }

                            newRowIndex++;
                            copiedRowsCount++;
                        }
                    }
                }

                // Автоподбор ширины колонок
                newSheet.Cells[newSheet.Dimension.Address].AutoFitColumns();

                // 5. Сохранение
                package.Save();

                Console.WriteLine($"Копирование завершено. Скопировано строк: {copiedRowsCount}");
            }
        }
    }

    public partial class Form1 : Form
    {
        private ExcelPackage? _excelPackage;      // Основной файл
        private ExcelPackage? _auxExcelPackage;
        private static readonly string[] COLS_OUT = {
            "Номер", "Наименование", "Приоритет"
        };
        private DataGridView mainDgv;
        private Dictionary<string, int> _visioPriorityMap = new Dictionary<string, int>();
        private List<Dictionary<string, string>> data;
        private DataGridView dataGridView = null!;
        private Label lblFileInfo = null!;
        private Label lblStatus = null!;
        private Panel statusPanel = null!;
        private string _visioSourceSheetName = "1.ТЗ на объект ZONT для VISIO";
        // Коэффициент для перевода из миллиметров в дюймы (1 дюйм = 25.4 мм)
        private const double MM_TO_INCH = 1.0 / 25.4;
        // НОВОЕ ПОЛЕ для хранения детальной, сырой информации
        private List<RawExcelHit> _rawHits = new List<RawExcelHit>();
        //private TextBox _textBoxTargetSheets = null!;
        //private TextBox _textBoxSearchWords = null!;

        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            AppSettings.Load();
            this.Icon = ZontSpecExtractor.Properties.Resources.picsart;
            SetupUI();
            data = new List<Dictionary<string, string>>();
            UpdateStatus($"Настройки загружены.");
        }
        private int ExcelColumnLetterToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return 0;
            columnName = columnName.ToUpperInvariant();
            int sum = 0;
            foreach (char c in columnName)
            {
                if (c < 'A' || c > 'Z') return 0; // Некорректный символ
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }

        public class VisioDrawItem
        {
            public string MasterName { get; set; } = "";
            public string Text { get; set; } = ""; // Текст для подписи фигуры (если нужно)

            // Тип размещения: "Manual" (фиксированные координаты) или "Sequential" (змейка)
            public string PlacementType { get; set; } = "Sequential";

            // Координаты (в миллиметрах)
            public double X_MM { get; set; } = 0;
            public double Y_MM { get; set; } = 0;

            // Привязка (Anchor)
            public string Anchor { get; set; } = "Center";
        }

        private Dictionary<string, EquipmentItem> _equipmentConfig = new Dictionary<string, EquipmentItem>(StringComparer.OrdinalIgnoreCase)
        {
            // Пример заполнения (это можно вынести в JSON или отдельный config-файл позже)
            { "Реле 12/220 16A", new EquipmentItem { ShortName = "Реле 12В", PositionCode = "K1", ShapeMasterName = "Relay_12V" } },
            { "Автомат 16А", new EquipmentItem { ShortName = "Авт. 16А", PositionCode = "QF1", ShapeMasterName = "CircuitBreaker" } },
            // Добавьте сюда остальные позиции из вашего ТЗ
        };

        public class EquipmentItem
        {
            public string OriginalName { get; set; } // Имя из Excel (для поиска)
            public string ShortName { get; set; }    // Короткое имя для схемы (из настроек)
            public string PositionCode { get; set; } // Код позиции (например, "QF1", "K1")
            public int Quantity { get; set; }        // Количество
            public string ShapeMasterName { get; set; } // Имя мастера в стенсиле Visio

            // Дополнительные данные, если нужно (ток, напряжение и т.д.)
        }

        private string GetColumnName(int columnIndex)
        {
            string dividend = string.Empty;
            string columnName = string.Empty;
            int modulo;

            while (columnIndex > 0)
            {
                modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                columnIndex = (int)((columnIndex - modulo) / 26);
            }
            return columnName;
        }

        private void SetupVisioPage(Visio.Page page, VisioConfiguration cfg)
        {
            Visio.Cell orientationCell = null;
            Visio.Cell widthCell = null;
            Visio.Cell heightCell = null;

            // Константы для ориентации (числовые значения VisPrintPageOrientation):
            // visPortrait = 1
            // visLandscape = 2
            const int visPortraitValue = 1;
            const int visLandscapeValue = 2;

            try
            {
                // 1. Получаем ячейки для флага ориентации и размеров страницы
                orientationCell = page.PageSheet.get_CellsU("PrintPageOrientation");
                widthCell = page.PageSheet.get_CellsU("PageWidth");
                heightCell = page.PageSheet.get_CellsU("PageHeight");

                // 2. Читаем текущие размеры в Дюймах (Visio Internal Units)
                // Если Visio изначально использует метрическую систему, он сам конвертирует эти значения.
                double currentWidth = widthCell.ResultIU;
                double currentHeight = heightCell.ResultIU;

                // 3. Устанавливаем новый флаг ориентации и корректируем размеры
                if (cfg.PageOrientation == "Landscape")
                {
                    // Устанавливаем флаг "Альбомная"
                    orientationCell.FormulaForce = visLandscapeValue.ToString();

                    // Если текущая ширина меньше текущей высоты (страница в портрете), меняем их местами
                    if (currentWidth < currentHeight)
                    {
                        // Меняем местами ширину и высоту (запись через ResultIU)
                        widthCell.ResultIU = currentHeight;
                        heightCell.ResultIU = currentWidth;
                    }
                }
                else if (cfg.PageOrientation == "Portrait")
                {
                    // Устанавливаем флаг "Книжная"
                    orientationCell.FormulaForce = visPortraitValue.ToString();

                    // Если текущая ширина больше текущей высоты (страница в альбомной), меняем их местами
                    if (currentWidth > currentHeight)
                    {
                        // Меняем местами ширину и высоту
                        widthCell.ResultIU = currentHeight;
                        heightCell.ResultIU = currentWidth;
                    }
                }
            }
            finally
            {
                // Обязательное освобождение COM-объектов
                ReleaseComObject(orientationCell);
                ReleaseComObject(widthCell);
                ReleaseComObject(heightCell);
            }
        }

        private bool SetShapeData(Visio.Shape shape, string labelToFind, string newValue)
        {
            if (shape == null || string.IsNullOrWhiteSpace(newValue)) return false;

            try
            {
                // Проверяем наличие секции свойств
                if (shape.get_SectionExists((short)Visio.VisSectionIndices.visSectionProp, (short)Visio.VisExistsFlags.visExistsAnywhere) == 0)
                    return false;

                short propSection = (short)Visio.VisSectionIndices.visSectionProp;
                short rowCount = shape.get_RowCount(propSection);

                for (short row = 0; row < rowCount; row++)
                {
                    // Получаем метку (Label)
                    Visio.Cell labelCell = shape.get_CellsSRC(propSection, row, (short)Visio.VisCellIndices.visCustPropsLabel);
                    string label = labelCell.get_ResultStr("");

                    // Сравниваем название поля
                    if (string.Equals(label.Trim(), labelToFind.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        // Формируем значение для записи
                        string formulaVal = "\"" + newValue.Replace("\"", "\"\"") + "\"";

                        Visio.Cell valueCell = shape.get_CellsSRC(propSection, row, (short)Visio.VisCellIndices.visCustPropsValue);
                        valueCell.FormulaU = formulaVal;

                        return true; // УСПЕХ: Поле найдено и обновлено
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка записи данных фигуры: {ex.Message}");
            }

            return false; // Поле не найдено
        }

        // ПРИМЕЧАНИЕ: Вызовите этот метод (SetupVisioPage) **сразу** после создания страницы Visio.
        // Например, если страница создается в другом месте:
        // Visio.Page page = newDoc.Pages.Add();
        // SetupVisioPage(page, config); // <-- Вызов здесь

        private string SanitizeFileName(string fileName)
        {
            // Удаляем недопустимые символы для имени файла Windows
            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            // Экранируем символы и заменяем их на подчеркивание
            string invalidRe = string.Format(@"[{0}]", System.Text.RegularExpressions.Regex.Escape(new string(invalidChars)));
            return System.Text.RegularExpressions.Regex.Replace(fileName, invalidRe, "_");
        }

        private Visio.Master GetMasterByName(Visio.Document stencilDoc, string masterName)
        {
            try
            {
                // Trim() для удаления лишних пробелов по краям, ToLower() для нормализации.
                string cleanedName = masterName.Trim();

                // Попробуем найти Master по имени. 
                // Если имя не найдено, Visio.Masters.Item() выбросит исключение.
                // Используем ItemU, если Master может иметь локализованное имя.
                return stencilDoc.Masters.get_ItemU(cleanedName);
            }
            catch (Exception)
            {
                // Ошибка - Master не найден.
                // Вы можете добавить сюда логирование или уведомление.
                // System.Diagnostics.Debug.WriteLine($"Master '{masterName}' не найден.");
                return null;
            }
        }

        // Хелпер для конвертации "A" -> 1, "AA" -> 27
        private int GetColumnIndex(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return 0;

            // Если пользователь ввел число
            if (int.TryParse(columnName, out int index)) return index;

            // Если буквы
            columnName = columnName.ToUpperInvariant();
            int sum = 0;
            foreach (char c in columnName)
            {
                if (c < 'A' || c > 'Z') return 0; // Невалидный символ
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }

        // Добавить в класс Form1
        private void UpdateConfigurationsWithRealSizes()
        {
            // 1. Собираем все пути к трафаретам
            var allStencils = AppSettings.SchemeConfig.StencilFilePaths
                .Union(AppSettings.LabelingConfig.StencilFilePaths)
                .Union(AppSettings.CabinetConfig.StencilFilePaths)
                .Distinct()
                .ToList();

            // 2. Собираем все имена мастеров, которые используются в конфигах
            var allMasters = new List<string>();

            // Из фиксированных фигур
            allMasters.AddRange(AppSettings.SchemeConfig.PredefinedMasterConfigs.Select(x => x.MasterName));
            allMasters.AddRange(AppSettings.LabelingConfig.PredefinedMasterConfigs.Select(x => x.MasterName));
            allMasters.AddRange(AppSettings.CabinetConfig.PredefinedMasterConfigs.Select(x => x.MasterName));

            // Из правил поиска (найденные фигуры)
            allMasters.AddRange(AppSettings.SchemeConfig.SearchRules.Select(x => x.VisioMasterName));
            allMasters.AddRange(AppSettings.LabelingConfig.SearchRules.Select(x => x.VisioMasterName));
            allMasters.AddRange(AppSettings.CabinetConfig.SearchRules.Select(x => x.VisioMasterName));

            allMasters = allMasters.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToList();

            if (allMasters.Count == 0) return;

            UpdateStatus("? Обновление размеров фигур из Visio...");

            // 3. ПАКЕТНОЕ ПОЛУЧЕНИЕ РАЗМЕРОВ (Один вызов Visio)
            var sizes = VisioScanner.GetMastersDimensionsBatch(allMasters, allStencils);

            // 4. Применяем размеры ко всем конфигам
            ApplySizesToConfig(AppSettings.SchemeConfig, sizes);
            ApplySizesToConfig(AppSettings.LabelingConfig, sizes);
            ApplySizesToConfig(AppSettings.CabinetConfig, sizes);

            UpdateStatus("? Размеры фигур обновлены.");
        }

        private void ApplySizesToConfig(VisioConfiguration cfg, Dictionary<string, (double w, double h)> sizes)
        {
            if (cfg.PredefinedMasterConfigs == null) return;

            foreach (var pm in cfg.PredefinedMasterConfigs)
            {
                // ИСПРАВЛЕНИЕ: Проверка на null
                if (string.IsNullOrWhiteSpace(pm.MasterName)) continue;

                if (sizes.TryGetValue(pm.MasterName, out var dim))
                {
                    pm.Width = dim.w;
                    pm.Height = dim.h;
                }
            }
        }

        // В обработчике кнопки настроек ОБЯЗАТЕЛЬНО вызовите этот метод
        private void OpenVisioSettingsClick(object? sender, EventArgs e)
        {
            // Сначала обновляем размеры, чтобы превью не было "красными точками"
            UpdateConfigurationsWithRealSizes();

            using (var settingsForm = new GeneralSettingsForm())
            {
                if (settingsForm.ShowDialog(this) == DialogResult.OK)
                {
                    UpdateStatus("Общие настройки сохранены.");
                }
            }
        }


        /// <summary>
        /// Определяет размеры страницы в дюймах на основе формата и ориентации.
        /// </summary>
        private (double width, double height) GetPageDimensions(string size, string orientation)
        {
            // Размеры А4 в дюймах (210x297 мм)
            const double A4_WIDTH = 8.2677;
            const double A4_HEIGHT = 11.6929;

            // В будущем здесь можно добавить логику для А3 и других форматов
            double w = A4_WIDTH;
            double h = A4_HEIGHT;

            // Если ориентация Landscape, меняем местами ширину и высоту
            if (orientation.Equals("Landscape", StringComparison.OrdinalIgnoreCase))
            {
                return (h, w);
            }

            return (w, h); // Portrait
        }

        // =========================================================================
        // НОВЫЙ КОМПОНЕНТ: Панель предпросмотра листа и фигур
        // =========================================================================
        // Внутри класса Form1 или отдельным классом
        // Разместить где-то в Form1.cs (например, в конце класса Form1)
        public class PagePreviewPanel : Panel
        {
            public string PageSize { get; set; } = "A4";
            public string Orientation { get; set; } = "Portrait";

            // Фиксированные фигуры (было 'Shapes', стало 'FixedShapes')
            public List<PredefinedMasterConfig> FixedShapes { get; set; } = new List<PredefinedMasterConfig>();

            // Настройки для потока (Sequential)
            public SequentialDrawingConfig SeqConfig { get; set; }

            public PagePreviewPanel()
            {
                this.DoubleBuffered = true;
                this.BackColor = Color.LightGray;
                this.Paint += OnPaint;
            }

            private (double w, double h) GetPageSizeMM()
            {
                double w = 210, h = 297;
                if (PageSize == "A3") { w = 297; h = 420; }
                if (Orientation == "Landscape") (w, h) = (h, w);
                return (w, h);
            }

            private void OnPaint(object? sender, PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var (pageW, pageH) = GetPageSizeMM();
                float margin = 20;
                float scale = Math.Min((this.Width - margin * 2) / (float)pageW, (this.Height - margin * 2) / (float)pageH);

                float drawW = (float)pageW * scale;
                float drawH = (float)pageH * scale;
                float startX = (this.Width - drawW) / 2;
                float startY = (this.Height - drawH) / 2;

                // Рисуем лист
                g.FillRectangle(Brushes.White, startX, startY, drawW, drawH);
                g.DrawRectangle(Pens.Black, startX, startY, drawW, drawH);
                g.DrawString("Низ листа (Visio 0,0)", this.Font, Brushes.Gray, startX, startY + drawH + 5);

                // Функция перевода координат MM -> Screen Pixels
                PointF ToScreen(double xMM, double yMM)
                {
                    return new PointF(
                        startX + (float)xMM * scale,
                        startY + drawH - (float)yMM * scale
                    );
                }

                // 1. РИСУЕМ ФИКСИРОВАННЫЕ ФИГУРЫ (СИНИЕ)
                if (FixedShapes != null)
                {
                    using (var brush = new SolidBrush(Color.FromArgb(80, 0, 0, 255)))
                    using (var pen = new Pen(Color.Blue, 1))
                    {
                        foreach (var shape in FixedShapes)
                        {
                            if (string.IsNullOrWhiteSpace(shape.CoordinatesXY)) continue;
                            var parts = shape.CoordinatesXY.Split(',');

                            // === БЕЗОПАСНЫЙ ПАРСИНГ (Ваш фикс) ===
                            float xVal = 0, yVal = 0;
                            bool xOk = parts.Length >= 1 && float.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out xVal);
                            bool yOk = parts.Length >= 2 && float.TryParse(parts[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out yVal);

                            if (!xOk || !yOk) continue;
                            // =====================================

                            // Размеры (дефолт если < 1)
                            float w = (float)(shape.Width < 1 ? 30 : shape.Width);
                            float h = (float)(shape.Height < 1 ? 15 : shape.Height);

                            // Переводим координаты MM -> Пиксели экрана
                            var pt = ToScreen(xVal, yVal);

                            float rectX, rectY;
                            // Нормализуем Anchor (убираем пробелы, null check)
                            string anchor = (shape.Anchor ?? "Center").Trim();

                            // === ЛОГИКА ОТРИСОВКИ (ИСПРАВЛЕННАЯ) ===
                            if (string.Equals(anchor, "TopLeft", StringComparison.OrdinalIgnoreCase))
                            {
                                // Точка привязки = Левый ВЕРХНИЙ угол.
                                // В GDI+ Y растет вниз, поэтому рисуем прямо от точки (pt).
                                rectX = pt.X;
                                rectY = pt.Y;
                            }
                            else if (string.Equals(anchor, "BottomLeft", StringComparison.OrdinalIgnoreCase))
                            {
                                // Точка привязки = Левый НИЖНИЙ угол.
                                // Чтобы прямоугольник стоял НАД точкой, нужно ВЫЧЕСТЬ высоту (подняться вверх по экрану).
                                rectX = pt.X;
                                rectY = pt.Y - (h * scale);
                            }
                            else // Center и любые другие варианты
                            {
                                // Точка привязки = ЦЕНТР фигуры.
                                rectX = pt.X - (w * scale) / 2;
                                rectY = pt.Y - (h * scale) / 2;
                            }
                            // ========================================

                            var rect = new RectangleF(rectX, rectY, w * scale, h * scale);

                            // Рисуем полупрозрачный фон и рамку
                            g.FillRectangle(brush, rect);
                            g.DrawRectangle(pen, rect.X, rect.Y, rect.Width, rect.Height);

                            // Рисуем красную точку привязки (поверх всего, чтобы проверить точность)
                            g.FillEllipse(Brushes.Red, pt.X - 3, pt.Y - 3, 6, 6);

                            // Подпись имени мастера
                            g.DrawString(shape.MasterName, SystemFonts.DefaultFont, Brushes.Black, rect.X, rect.Y - 12);
                        }
                    }
                }

                // 2. РИСУЕМ ПОТОК НАЙДЕННЫХ ФИГУР (ЗЕЛЕНЫЕ "ФАНТОМЫ")
                if (SeqConfig != null && SeqConfig.Enabled)
                {
                    var parts = SeqConfig.StartCoordinatesXY.Split(',');
                    if (parts.Length >= 2)
                    {
                        double curX = double.Parse(parts[0], CultureInfo.InvariantCulture);
                        double curY = double.Parse(parts[1], CultureInfo.InvariantCulture);
                        double maxW = SeqConfig.MaxLineWidthMM;
                        double hGap = SeqConfig.HorizontalStepMM;
                        double vGap = SeqConfig.VerticalStepMM;

                        using (var brush = new SolidBrush(Color.FromArgb(80, 0, 255, 0)))
                        using (var pen = new Pen(Color.Green, 1) { DashStyle = DashStyle.Dash })
                        {
                            double startLineX = curX;
                            double testW = 40; // Дефолтная ширина фантома (как вы просили)
                            double testH = 20; // Дефолтная высота фантома
                            double rowMaxH = 0; // Для симуляции

                            for (int i = 1; i <= 6; i++)
                            {
                                // Проверка переноса строки (симуляция)
                                if (curX + testW - startLineX > maxW)
                                {
                                    curX = startLineX;
                                    curY = curY - rowMaxH - vGap;
                                    rowMaxH = 0;
                                }

                                // У нас расчет идет от TopLeft, поэтому точка привязки = TopLeft фигуры
                                var screenPt = ToScreen(curX, curY);

                                var rect = new RectangleF(screenPt.X, screenPt.Y, (float)testW * scale, (float)testH * scale);

                                g.FillRectangle(brush, rect);
                                g.DrawRectangle(pen, rect.X, rect.Y, rect.Width, rect.Height);
                                g.DrawString($"#{i} (Поток)", SystemFonts.DefaultFont, Brushes.DarkGreen, rect.X, rect.Y);

                                // Обновление курсора
                                curX += testW + hGap;
                                if (testH > rowMaxH) rowMaxH = testH;
                            }

                            // Рисуем линию ограничения ширины
                            var limitPt = ToScreen(startLineX + maxW, 0);
                            float limitX = limitPt.X;
                            g.DrawLine(Pens.Red, limitX, startY, limitX, startY + drawH);
                            g.DrawString("Max Width", SystemFonts.DefaultFont, Brushes.Red, limitX + 2, startY + 10);
                        }
                    }
                }
            }
        }


        private void UpdateStatus(string message)
        {
            if (this.lblStatus == null) return;

            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateStatus(message)));
            }
            else
            {
                this.lblStatus.Text = message;
            }
        }

        private void SetupUI()
        {
            this.Text = "ZONT Spec Extractor";
            this.Size = new System.Drawing.Size(1200, 650);
            this.MinimumSize = new System.Drawing.Size(900, 450);
            this.BackColor = System.Drawing.Color.White;
            this.Font = new System.Drawing.Font("Segoe UI", 7F);
            this.StartPosition = FormStartPosition.CenterScreen;

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                RowStyles =
                {
                    new RowStyle(SizeType.Absolute, 50),
                    new RowStyle(SizeType.Percent, 100),
                    new RowStyle(SizeType.Absolute, 25)
                },
                Padding = new Padding(0)
            };

            var headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                Padding = new Padding(10, 5, 10, 5)
            };

            var buttonFlowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(0)
            };

            lblFileInfo = new Label
            {
                Text = "Файл не выбран",
                Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.Black,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Left,
                Margin = new Padding(0, 0, 0, 0),
                AutoSize = true
            };

            const int ICON_SIZE = 34;

            // Переименована кнопка для соответствия новому обработчику
            var btnLoad = CreateStyledButton("Загрузить Excel", System.Drawing.Color.FromArgb(46, 139, 87), System.Drawing.Color.White);
            btnLoad.Click += BtnLoadFile_Click;

            System.Drawing.Image originalExcelImage = Properties.Resources.iconexcel.ToBitmap();

            // Масштабируем изображение перед присвоением
            btnLoad.Image = originalExcelImage.GetThumbnailImage(ICON_SIZE, ICON_SIZE, null, IntPtr.Zero);
            btnLoad.ImageAlign = ContentAlignment.MiddleLeft;    // Иконка вверху
            btnLoad.TextAlign = ContentAlignment.MiddleCenter;  // Текст внизу
            btnLoad.Height = 35; // Увеличьте высоту кнопки, чтобы вместить иконку и текст
                                 // btnLoad.Font = new Font(btnLoad.Font.FontFamily, 8); // Опционально: уменьшите шрифт

            var btnCreateTable = CreateStyledButton("Создать таблицу", System.Drawing.Color.FromArgb(46, 204, 113), System.Drawing.Color.White);
            btnCreateTable.Click += CreateTableClick;
            btnCreateTable.Image = originalExcelImage.GetThumbnailImage(ICON_SIZE, ICON_SIZE, null, IntPtr.Zero);
            btnCreateTable.ImageAlign = ContentAlignment.MiddleLeft;
            btnCreateTable.TextAlign = ContentAlignment.MiddleCenter;
            btnCreateTable.Height = 35;

            // --- Кнопка СГЕНЕРИРОВАТЬ VISIO ---
            var btnOpenVisio = CreateStyledButton("Сгенерировать Visio", System.Drawing.Color.FromArgb(21, 96, 189), System.Drawing.Color.White);
            btnOpenVisio.Click += OpenVisioClick;

            System.Drawing.Image originalVisioImage = Properties.Resources.iconvisio.ToBitmap();

            // Масштабируем изображение перед присвоением
            btnOpenVisio.Image = originalVisioImage.GetThumbnailImage(ICON_SIZE, ICON_SIZE, null, IntPtr.Zero);
            btnOpenVisio.ImageAlign = ContentAlignment.MiddleLeft;   // Иконка вверху
            btnOpenVisio.TextAlign = ContentAlignment.MiddleCenter; // Текст внизу
            btnOpenVisio.Height = 35; // Увеличьте высоту кнопки
                                      // btnOpenVisio.Font = new Font(btnOpenVisio.Font.FontFamily, 8); // Опционально: уменьшите шрифт

            var btnVisioSettings = CreateStyledButton("Настройки", System.Drawing.Color.FromArgb(245, 245, 245), System.Drawing.Color.Black);
            btnVisioSettings.Click += OpenVisioSettingsClick;

            System.Drawing.Image originalSettingsImage = Properties.Resources.icongear.ToBitmap();

            btnVisioSettings.Image = originalSettingsImage.GetThumbnailImage(ICON_SIZE, ICON_SIZE, null, IntPtr.Zero);
            btnVisioSettings.ImageAlign = ContentAlignment.MiddleLeft;    // Иконка вверху
            btnVisioSettings.TextAlign = ContentAlignment.MiddleCenter;  // Текст внизу
            btnVisioSettings.Height = 35; // Увеличьте высоту кнопки, чтобы вместить иконку и текст
                                          // btnLoad.Font = new Font(btnLoad.Font.FontFamily, 8); // Опционально: уменьшите шрифт

            buttonFlowPanel.Controls.Add(btnLoad);
            buttonFlowPanel.Controls.Add(btnCreateTable);
            buttonFlowPanel.Controls.Add(btnOpenVisio);
            buttonFlowPanel.Controls.Add(btnVisioSettings);


            headerPanel.Controls.Add(buttonFlowPanel);
            headerPanel.Controls.Add(lblFileInfo);

            // Настройка DataGridView
            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                Font = new System.Drawing.Font("Segoe UI", 9F),
                GridColor = System.Drawing.Color.LightGray
            };

            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(60, 60, 60);
            dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dataGridView.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            dataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dataGridView.ColumnHeadersHeight = 35;

            dataGridView.RowHeadersVisible = false;
            dataGridView.RowTemplate.Height = 28;
            dataGridView.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 245, 245);
            dataGridView.DefaultCellStyle.BackColor = System.Drawing.Color.White;
            dataGridView.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(180, 215, 255);
            dataGridView.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;

            foreach (var colName in COLS_OUT)
            {
                int width = 120;
                if (colName == "Наименование") width = 300;
                else if (colName == "Лист") width = 220;
                else if (colName == "Цена с НДС, руб." || colName == "Сумма с НДС, руб.") width = 150;

                dataGridView.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = colName,
                    HeaderText = colName,
                    Width = width,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Padding = new Padding(8, 0, 8, 0),
                        Alignment = colName == "Наименование" ? DataGridViewContentAlignment.MiddleLeft : DataGridViewContentAlignment.MiddleCenter
                    }
                });
            }

            statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 25,
                BackColor = System.Drawing.Color.FromArgb(220, 220, 220),
                Padding = new Padding(10, 0, 10, 0)
            };

            lblStatus = new Label
            {
                Text = "",
                Font = new System.Drawing.Font("Segoe UI", 8F),
                ForeColor = System.Drawing.Color.Black,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true
            };
            statusPanel.Controls.Add(lblStatus);

            mainLayout.Controls.Add(headerPanel, 0, 0);
            mainLayout.Controls.Add(dataGridView, 0, 1);
            mainLayout.Controls.Add(statusPanel, 0, 2);

            this.Controls.Add(mainLayout);
        }

        // --- 4. ОСНОВНАЯ ФОРМА: Вывод найденных фигур ---
        /// <summary>
        /// Выводит список найденных мастеров Visio в строку статуса.
        /// </summary>
        // ИСПРАВЛЕНИЕ: Метод теперь принимает два объекта VisioConfiguration
        private void ShowDiscoveredMasters(VisioConfiguration configMarking, VisioConfiguration configScheme)
        {
            var allMasters = configMarking.AvailableMasters
                .Union(configScheme.AvailableMasters)
                .Distinct()
                .ToList();

            if (allMasters.Any())
            {
                // Ограничиваем список для компактного вывода в строку статуса
                string masterList = string.Join(", ", allMasters.OrderBy(m => m).Take(10).ToArray());
                if (allMasters.Count > 10)
                {
                    masterList += $" и еще {allMasters.Count - 10} фигур...";
                }

                UpdateStatus($"Найденные фигуры: {masterList} | Всего: {allMasters.Count}");
            }
            else
            {
                UpdateStatus("?? Не удалось найти ни одной фигуры Visio в указанных трафаретах.");
            }
        }

        private Button CreateStyledButton(string text, System.Drawing.Color backgroundColor, System.Drawing.Color foregroundColor)
        {
            var button = new Button
            {
                Text = text,
                Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold),
                Size = new System.Drawing.Size(180, 40),
                BackColor = backgroundColor,
                ForeColor = foregroundColor,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Margin = new Padding(5, 0, 5, 0)
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(
                (int)(backgroundColor.R * 0.9),
                (int)(backgroundColor.G * 0.9),
                (int)(backgroundColor.B * 0.9)
            );
            button.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(
                (int)(backgroundColor.R * 0.7),
                (int)(backgroundColor.G * 0.7),
                (int)(backgroundColor.B * 0.7)
            );
            return button;
        }

        // --- НОВЫЙ ВСПОМОГАТЕЛЬНЫЙ МЕТОД: Получение имен листов ---
        private List<string> GetSheetNames(string filePath)
        {
            var sheetNames = new List<string>();
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    sheetNames.AddRange(package.Workbook.Worksheets.Select(ws => ws.Name));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении имен листов: {ex.Message}", "Ошибка Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine($"Ошибка при чтении имен листов: {ex.Message}");
            }
            return sheetNames;
        }


        // --- ОБНОВЛЕННЫЙ ОБРАБОТЧИК: Открытие файла и выбор листов ---
        private void BtnLoadFile_Click(object? sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Выберите .xlsx файл";
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    AppSettings.LastLoadedFilePath = openFileDialog.FileName;

                    // Шаг 1: Получаем все имена листов
                    var allSheetNames = GetSheetNames(AppSettings.LastLoadedFilePath);

                    if (!allSheetNames.Any())
                    {
                        MessageBox.Show("Не удалось получить имена листов. Проверьте файл.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Шаг 2: Диалог - Выбор листов для АНАЛИЗА
                    // Используем ранее сохраненный список для предвыбора (если есть)
                    using (var sheetForm = new SheetSelectionForm(allSheetNames, AppSettings.SearchConfig.TargetSheetNames))
                    {
                        sheetForm.Text = "Выберите листы для АНАЛИЗА (Поиск позиций)"; // Убрал нумерацию "1.", так как шаг теперь один

                        if (sheetForm.ShowDialog() == DialogResult.OK)
                        {
                            // Сохраняем настройки анализа
                            AppSettings.SearchConfig.TargetSheetNames = sheetForm.SelectedSheets;
                            AppSettings.Save();

                            // Сразу запускаем основную обработку, пропуская второй диалог
                            LoadFiles(new[] { AppSettings.LastLoadedFilePath });
                        }
                        else
                        {
                            UpdateStatus("Загрузка отменена пользователем.");
                        }
                    }
                }
            }
        }

        private void CreateTableClick(object? sender, EventArgs e)
        {
            if (!_rawHits.Any())
            {
                MessageBox.Show("Сначала необходимо загрузить и проанализировать Excel-файл!");
                return;
            }

            // =================================================================================
            // 1. ПОДГОТОВКА ДАННЫХ (Строгая привязка к строкам Excel)
            // =================================================================================

            // Мы группируем находки по Имени листа и Номеру строки.
            // Это гарантирует, что данные одной строки останутся вместе.
            var rowsData = _rawHits
                .Where(h => h.ConditionMet) // Берем только то, где выполнилось условие (где была "1")
                .GroupBy(h => new { h.SheetName, h.RowIndex }) // ГРУППИРОВКА ПО СТРОКЕ
                .Select(g => new
                {
                    // Сортировка: сначала по листам, потом по номеру строки в Excel
                    SortSheet = g.Key.SheetName,
                    SortRow = g.Key.RowIndex,

                    // Ищем данные для каждой колонки ИМЕННО В ЭТОЙ СТРОКЕ (g)

                    // Основная позиция (TargetColumn = "Position" или пусто)
                    Position = g.FirstOrDefault(x => string.IsNullOrEmpty(x.TargetColumn) || x.TargetColumn == "Position")?.FullItemName,

                    // Количество (берем от позиции или 1)
                    Quantity = g.Where(x => string.IsNullOrEmpty(x.TargetColumn) || x.TargetColumn == "Position").Sum(x => x.Quantity),

                    // Остальные колонки. Если в этой строке нет данных для Col4, будет null (пустота)
                    Col4 = g.FirstOrDefault(x => x.TargetColumn == "Col4")?.FullItemName,
                    Col5 = g.FirstOrDefault(x => x.TargetColumn == "Col5")?.FullItemName,
                    Col6 = g.FirstOrDefault(x => x.TargetColumn == "Col6")?.FullItemName,
                    Col7 = g.FirstOrDefault(x => x.TargetColumn == "Col7")?.FullItemName,
                    Col8 = g.FirstOrDefault(x => x.TargetColumn == "Col8")?.FullItemName,
                    Col9 = g.FirstOrDefault(x => x.TargetColumn == "Col9")?.FullItemName,
                    Col10 = g.FirstOrDefault(x => x.TargetColumn == "Col10")?.FullItemName,
                    Col11 = g.FirstOrDefault(x => x.TargetColumn == "Col11")?.FullItemName,
                    Col12 = g.FirstOrDefault(x => x.TargetColumn == "Col12")?.FullItemName,
                    Col13 = g.FirstOrDefault(x => x.TargetColumn == "Col13")?.FullItemName,
                    Col14 = g.FirstOrDefault(x => x.TargetColumn == "Col14")?.FullItemName,
                    Col15 = g.FirstOrDefault(x => x.TargetColumn == "Col15")?.FullItemName,
                    Col16 = g.FirstOrDefault(x => x.TargetColumn == "Col16")?.FullItemName,
                    Col17 = g.FirstOrDefault(x => x.TargetColumn == "Col17")?.FullItemName,
                    Col18 = g.FirstOrDefault(x => x.TargetColumn == "Col18")?.FullItemName
                })
                // СОРТИРОВКА: Строго как в Excel (по порядку строк)
                .OrderBy(x => x.SortSheet).ThenBy(x => x.SortRow)
                .ToList();

            if (!rowsData.Any())
            {
                MessageBox.Show("Не найдено данных для формирования таблицы.");
                return;
            }

            // =================================================================================
            // 2. СОЗДАНИЕ И НАСТРОЙКА ТАБЛИЦЫ
            // =================================================================================

            // Безопасное получение имени листа (исправление ошибки CS0029)
            string visioSheetName = "";
            try
            {
                var source = AppSettings.SearchConfig.VisioSourceSheetName;
                if (source is IEnumerable<string> list) visioSheetName = list.FirstOrDefault() ?? "";
                else visioSheetName = source?.ToString() ?? "";
            }
            catch { visioSheetName = "Ошибка имени"; }

            var tableForm = new Form
            {
                Text = $"Спецификация (Лист Visio: {(string.IsNullOrEmpty(visioSheetName) ? "Не выбран" : visioSheetName)})",
                Size = new Size(1400, 600),
                StartPosition = FormStartPosition.CenterParent
            };

            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };

            // --- ДОБАВЛЕНИЕ КОЛОНОК ---
            dgv.Columns.Add("Priority", "Номер приоритета");
            dgv.Columns.Add("Position", "Позиция (дословно)"); dgv.Columns["Position"].Width = 200;
            dgv.Columns.Add("Quantity", "Количество");

            dgv.Columns.Add("Col4", "Котёл/Насос/Кран");
            dgv.Columns.Add("Col5", "Кл. Насос/Котёл");
            dgv.Columns.Add("Col6", "Смеситель/кран2/охрана");
            dgv.Columns.Add("Col7", "Кл. Смесит/Кран ОТКР");
            dgv.Columns.Add("Col8", "Кл. Смесит/Кран ЗАКР");
            dgv.Columns.Add("Col9", "Кл. 1 Шина управл");
            dgv.Columns.Add("Col10", "Кл. 2 Шина управл");
            dgv.Columns.Add("Col11", "Минус");
            dgv.Columns.Add("Col12", "Датчики 1");
            dgv.Columns.Add("Col13", "Датчик2 / Питание +5В");
            dgv.Columns.Add("Col14", "Кл. 1 Датчик (-12В)");
            dgv.Columns.Add("Col15", "Кл. 2 Датчик T1");
            dgv.Columns.Add("Col16", "Кл. 3 Датчик T2");
            dgv.Columns.Add("Col17", "Кл. 4 Датчик воздух");
            dgv.Columns.Add("Col18", "Кл. 5 Датчик (+5В)");

            // =================================================================================
            // 3. ЗАПОЛНЕНИЕ ДАННЫМИ (Строго по строкам)
            // =================================================================================

            for (int i = 0; i < rowsData.Count; i++)
            {
                var dataItem = rowsData[i];

                int rowIndex = dgv.Rows.Add();
                var row = dgv.Rows[rowIndex];

                // Номер п/п (просто порядковый номер в итоговой таблице)
                row.Cells["Priority"].Value = i + 1;

                // Позиция
                row.Cells["Position"].Value = dataItem.Position;

                // Количество (если 0, можно не писать или писать 0 - по желанию)
                // Если позиция пустая, но есть насос, количество может быть 0, если не настроено иначе
                row.Cells["Quantity"].Value = dataItem.Quantity == 0 ? 1 : dataItem.Quantity;

                // Заполняем остальные колонки. 
                // Если в dataItem.Col4 ничего нет (null), в ячейку запишется null (пустота).
                // Следующее значение НЕ перепрыгнет сюда.
                row.Cells["Col4"].Value = dataItem.Col4;
                row.Cells["Col5"].Value = dataItem.Col5;
                row.Cells["Col6"].Value = dataItem.Col6;
                row.Cells["Col7"].Value = dataItem.Col7;
                row.Cells["Col8"].Value = dataItem.Col8;
                row.Cells["Col9"].Value = dataItem.Col9;
                row.Cells["Col10"].Value = dataItem.Col10;
                row.Cells["Col11"].Value = dataItem.Col11;
                row.Cells["Col12"].Value = dataItem.Col12;
                row.Cells["Col13"].Value = dataItem.Col13;
                row.Cells["Col14"].Value = dataItem.Col14;
                row.Cells["Col15"].Value = dataItem.Col15;
                row.Cells["Col16"].Value = dataItem.Col16;
                row.Cells["Col17"].Value = dataItem.Col17;
                row.Cells["Col18"].Value = dataItem.Col18;
            }

            tableForm.Controls.Add(dgv);
            tableForm.ShowDialog();
        }

        private async void LoadFiles(string[] filePaths)
        {
            // Очистка интерфейса перед стартом
            dataGridView.Rows.Clear();
            data.Clear();
            _rawHits.Clear();

            if (filePaths.Length > 0)
            {
                lblFileInfo.Text = $"📄 {System.IO.Path.GetFileName(filePaths[0])}";
            }
            else
            {
                lblFileInfo.Text = "Файл не выбран";
                return;
            }

            UpdateStatus("⏳ Идет сканирование Excel-файла...");

            try
            {
                await Task.Run(() =>
                {
                    foreach (var path in filePaths)
                    {
                        // Сканируем только те листы, которые выбрали в ПЕРВОМ диалоге 
                        // (ScanSpecificSheet берет их из AppSettings.SearchConfig.TargetSheetNames)
                        var hits = ScanSpecificSheet(path);
                        _rawHits.AddRange(hits);
                    }
                });

                // 1. Агрегируем сырые данные (_rawHits) в формат для основного просмотра
                data = GroupRawHitsForVisio(_rawHits);

                // 2. Обновляем таблицу на главной форме
                UpdateDataGridView();
                ShowResultMessage(_rawHits.Count);

                // Примечание: _visioSourceSheetName сейчас просто хранится в памяти 
                // и будет использован при нажатии кнопки "Создать таблицу".
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("Ошибка загрузки.");
            }

            var rowsData = _rawHits
        .Where(h => h.ConditionMet)
        .GroupBy(h => new { h.SheetName, h.RowIndex })
        .Select(g => new
        {
            SortSheet = g.Key.SheetName,
            SortRow = g.Key.RowIndex,
            Position = g.FirstOrDefault(x => string.IsNullOrEmpty(x.TargetColumn) || x.TargetColumn == "Position")?.FullItemName,
            Col4 = g.FirstOrDefault(x => x.TargetColumn == "Col4")?.FullItemName,
            Col5 = g.FirstOrDefault(x => x.TargetColumn == "Col5")?.FullItemName,
            Col6 = g.FirstOrDefault(x => x.TargetColumn == "Col6")?.FullItemName,
            Col7 = g.FirstOrDefault(x => x.TargetColumn == "Col7")?.FullItemName,
            Col12 = g.FirstOrDefault(x => x.TargetColumn == "Col12")?.FullItemName,
            Col14 = g.FirstOrDefault(x => x.TargetColumn == "Col14")?.FullItemName,
            Col13 = g.FirstOrDefault(x => x.TargetColumn == "Col13")?.FullItemName,
            Col15 = g.FirstOrDefault(x => x.TargetColumn == "Col15")?.FullItemName
        })
        .ToList();

            if (!rowsData.Any())
            {
                MessageBox.Show("Не найдено данных для формирования таблицы.");
                return;
            }

            // Функция парсинга приоритета
            double GetPriority(string? val)
            {
                if (string.IsNullOrEmpty(val)) return 999999.0;
                if (double.TryParse(val.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double res))
                    return res;
                return 999999.0;
            }

            // Сбор данных из разных колонок в один список
            var listFrom4 = rowsData.Where(x => !string.IsNullOrWhiteSpace(x.Col4)).Select(x => new
            { Name = x.Col4, PriorityRaw = x.Col5, PriorityVal = GetPriority(x.Col5), SourcePos = x.Position });
            var listFrom6 = rowsData.Where(x => !string.IsNullOrWhiteSpace(x.Col6)).Select(x => new
            { Name = x.Col6, PriorityRaw = x.Col7, PriorityVal = GetPriority(x.Col7), SourcePos = x.Position });
            var listFrom12 = rowsData.Where(x => !string.IsNullOrWhiteSpace(x.Col12)).Select(x => new
            { Name = x.Col12, PriorityRaw = x.Col14, PriorityVal = GetPriority(x.Col14), SourcePos = x.Position });
            var listFrom13 = rowsData.Where(x => !string.IsNullOrWhiteSpace(x.Col13)).Select(x => new
            { Name = x.Col13, PriorityRaw = x.Col15, PriorityVal = GetPriority(x.Col15), SourcePos = x.Position });

            var finalQueue = listFrom4
                .Concat(listFrom6)
                .Concat(listFrom12)
                .Concat(listFrom13)
                .OrderBy(x => x.PriorityVal)
                .ToList();

            // =================================================================================
            // 2. НАСТРОЙКА ТАБЛИЦЫ (ПОЛНАЯ ПЕРЕЗАГРУЗКА КОЛОНОК)
            // =================================================================================
            mainDgv = this.dataGridView;
            if (mainDgv != null)
            {
                mainDgv.AutoGenerateColumns = false;
                mainDgv.Columns.Clear();

                // --- ДОБАВИТЬ ЭТИ СТРОКИ ДЛЯ КОПИРОВАНИЯ ---
                // 1. Разрешаем копирование (EnableAlwaysIncludeHeaderText копирует и заголовки, удобно для Excel)
                mainDgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

                // 2. Разрешаем выделять несколько строк или ячеек
                mainDgv.MultiSelect = true;

                // 3. Режим выделения: RowHeaderSelect позволяет выделять и строки целиком (нажав слева), и отдельные ячейки
                mainDgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;        // Удаляем всё, что было раньше

                // Создаем 5 чистых колонок
                mainDgv.Columns.Add("BlockNum", "№");
                mainDgv.Columns.Add("Terminals", "Клеммы");
                mainDgv.Columns.Add("BlockName", "Наименование блока");
                mainDgv.Columns.Add("SourcePos", "Позиция из КП");
                mainDgv.Columns.Add("PriorityVal", "Приоритет");

                // Настройка ширины (для красоты)
                mainDgv.Columns["BlockNum"].Width = 40;
                mainDgv.Columns["Terminals"].Width = 80;
                mainDgv.Columns["BlockName"].Width = 250;
                mainDgv.Columns["SourcePos"].Width = 60;
                mainDgv.Columns["PriorityVal"].Width = 60;

                mainDgv.Rows.Clear();

                // =================================================================================
                // 3. ОСНОВНОЙ ЦИКЛ ЗАПОЛНЕНИЯ
                // =================================================================================

                // Поиск индекса ПОСЛЕДНЕГО Управл./Перекл. для вставки минуса внутри списка
                int lastSwitchIndex = finalQueue.FindLastIndex(x =>
                    (x.Name ?? "").Contains("Управл.", StringComparison.OrdinalIgnoreCase) ||
                    (x.Name ?? "").Contains("Перекл.", StringComparison.OrdinalIgnoreCase)
                );

                // Проверяем, есть ли вообще "Перекл." в списке (для финала таблицы)
                bool hasPereklForFooter = finalQueue.Any(x => (x.Name ?? "").Contains("Перекл.", StringComparison.OrdinalIgnoreCase));

                int currentTerminalCounter = 3; // Начинаем с 3-й клеммы
                int sensorsInGroup = 0;         // Счётчик текущей группы датчиков (0, 1, 2...)

                // Функция: это датчик?
                bool IsSensorRow(string name)
                {
                    if (string.IsNullOrEmpty(name)) return false;
                    return name.Contains("Т", StringComparison.Ordinal) ||
                           name.Contains("Дв", StringComparison.Ordinal);
                }

                string GetSmartBoilerName(string currentBlockName, string sourcePosName)
                {
                    if (string.IsNullOrWhiteSpace(sourcePosName) || string.IsNullOrWhiteSpace(currentBlockName))
                        return currentBlockName;

                    try
                    {
                        // Ищем паттерн: слово "котел", пробелы, цифра "1", пробелы, (тут модель), (реле)
                        // RegexOptions.IgnoreCase - чтобы не зависеть от регистра
                        var regex = new System.Text.RegularExpressions.Regex(
                            @"котел\s+1\s+(.*?)\(?реле\)?",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        var match = regex.Match(sourcePosName);

                        if (match.Success)
                        {
                            // Группа 1 - это то, что попало между "котел 1" и "реле"
                            string modelName = match.Groups[1].Value.Trim();

                            if (!string.IsNullOrEmpty(modelName))
                            {
                                // Заменяем цифру "1" как отдельное слово (\b1\b) на название модели
                                // Это предотвращает замену единиц внутри других чисел (например, 12В не пострадает)
                                return System.Text.RegularExpressions.Regex.Replace(currentBlockName, @"\b1\b", modelName);
                            }
                        }
                    }
                    catch { }

                    return currentBlockName;
                } 
           

                for (int i = 0; i < finalQueue.Count; i++)
                {
                    var item = finalQueue[i];

                    // 1. ПОЛУЧАЕМ "УМНОЕ" ИМЯ
                    // Если в исходной позиции (SourcePos) есть модель котла, она подставится вместо "1"
                    string originalName = item.Name ?? "";
                    string name = GetSmartBoilerName(originalName, item.SourcePos ?? "");

                    bool isSensor = IsSensorRow(name);

                    // --- 3.1 ВСТАВКА МИНУСА (ПЕРЕД ДАТЧИКАМИ) ---
                    if (isSensor)
                    {
                        if (sensorsInGroup == 0)
                        {
                            AddMinusRow(mainDgv, ref currentTerminalCounter);
                        }
                    }
                    else
                    {
                        sensorsInGroup = 0;
                    }

                    // --- 3.2 Добавляем основную строку ---
                    int r = mainDgv.Rows.Add();
                    var row = mainDgv.Rows[r];

                    // Расчет клемм (используем уже новое имя name)
                    string terminalValue = "";
                    if (name.Contains("Питание 220В", StringComparison.OrdinalIgnoreCase))
                    {
                        terminalValue = "1";
                    }
                    else if (name.Contains("Смеситель", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("Управл.", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("Перекл.", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("Д. давления", StringComparison.OrdinalIgnoreCase))
                    {
                        terminalValue = $"{currentTerminalCounter}, {currentTerminalCounter + 1}";
                        currentTerminalCounter += 2;
                    }
                    else
                    {
                        terminalValue = currentTerminalCounter.ToString();
                        currentTerminalCounter += 1;
                    }

                    // Заполнение
                    row.Cells["BlockNum"].Value = r + 1;
                    row.Cells["Terminals"].Value = terminalValue;

                    // ВАЖНО: Записываем в таблицу новое имя с Buderus
                    row.Cells["BlockName"].Value = name;

                    row.Cells["SourcePos"].Value = item.SourcePos;
                    row.Cells["PriorityVal"].Value = item.PriorityRaw;

                    // --- 3.3 УПРАВЛЕНИЕ ГРУППИРОВКОЙ ДАТЧИКОВ ---
                    if (isSensor)
                    {
                        sensorsInGroup++;
                        int sensorsAhead = 0;
                        // Для проверки следующих строк нам не нужно менять их имена, 
                        // достаточно проверить исходные IsSensorRow, так как "Т" или "Дв" не меняются при замене котла
                        for (int k = i + 1; k < finalQueue.Count; k++)
                        {
                            string nextName = finalQueue[k].Name ?? ""; // берем сырое имя для проверки типа
                            if (IsSensorRow(nextName)) sensorsAhead++;
                            else break;
                        }

                        if (sensorsInGroup == 2)
                        {
                            if (sensorsAhead != 1) sensorsInGroup = 0;
                        }
                        else if (sensorsInGroup >= 3)
                        {
                            sensorsInGroup = 0;
                        }
                    }
                }

                // =================================================================================
                // 4. "ПОДВАЛ" ТАБЛИЦЫ (ЕСЛИ ЕСТЬ "ПЕРЕКЛ.")
                // =================================================================================
                if (hasPereklForFooter)
                {
                    // 1. Минус -12В (GND)
                    AddRowManual(mainDgv, ref currentTerminalCounter, "Минус -12В (GND)", 1);

                    // 2. Питание +12В
                    AddRowManual(mainDgv, ref currentTerminalCounter, "Питание +12В", 1);

                    // 3. Шина связи RS485 (Занимает 2 клеммы!)
                    AddRowManual(mainDgv, ref currentTerminalCounter, "Шина связи RS485", 2);
                }

                mainDgv.Refresh();
            }
            else
            {
                MessageBox.Show("Не удалось найти главную таблицу.");
            }
        }

        // =================================================================================
        // ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
        // =================================================================================

        // 1. Метод для вставки обычного "автоматического" минуса
        private void AddMinusRow(DataGridView dgv, ref int terminalCounter)
        {
            int r = dgv.Rows.Add();
            var row = dgv.Rows[r];

            // Занимаем 1 клемму
            string term = terminalCounter.ToString();
            terminalCounter += 1;

            row.Cells["BlockNum"].Value = r + 1;
            row.Cells["Terminals"].Value = term;
            row.Cells["BlockName"].Value = "Минус -12В (GND)";
            row.Cells["SourcePos"].Value = "";
            row.Cells["PriorityVal"].Value = "";
        }

        // 2. Универсальный метод для добавления строки в конце (для RS485 и питания)
        private void AddRowManual(DataGridView dgv, ref int terminalCounter, string name, int terminalsCount)
        {
            int r = dgv.Rows.Add();
            var row = dgv.Rows[r];

            string termString = "";
            if (terminalsCount == 1)
            {
                termString = terminalCounter.ToString();
                terminalCounter += 1;
            }
            else // Если 2 клеммы
            {
                termString = $"{terminalCounter}, {terminalCounter + 1}";
                terminalCounter += 2;
            }

            row.Cells["BlockNum"].Value = r + 1;
            row.Cells["Terminals"].Value = termString;
            row.Cells["BlockName"].Value = name;
            row.Cells["SourcePos"].Value = "";
            row.Cells["PriorityVal"].Value = "";
        }

        // Вспомогательный метод для добавления строки с минусом (чтобы не дублировать код)

        private List<Dictionary<string, string>> GroupRawHitsForVisio(List<RawExcelHit> rawHits)
        {
            // Эта логика - это то, что было удалено из ScanSpecificSheet. 
            // Она группирует по ИСКОМОМУ СЛОВУ (SearchTerm), как требовалось для Visio ранее.

            return rawHits
                .GroupBy(h => h.SearchTerm) // Группируем по ИСКОМОМУ СЛОВУ
                .Select(g =>
                {
                    // Проверяем, было ли у этого правила ограничение (смотрим на флаг первой попавшейся записи группы)
                    bool isLimited = g.First().IsLimited;

                    int totalQty;
                    if (isLimited)
                    {
                        // Если стоит галочка "Ограничить", то независимо от количества найденных строк, сумма = 1
                        totalQty = 1;
                    }
                    else
                    {
                        // Иначе суммируем всё, что нашли
                        totalQty = g.Sum(x => x.Quantity);
                    }

                    return new Dictionary<string, string>
                    {
                        ["Лист"] = g.First().SheetName,
                        ["Наименование"] = g.Key,
                        ["Количество"] = totalQty.ToString()
                    };
                })
                .Where(x => x["Количество"] != "0")
                .ToList();
        }

        //private void btnApplySettings_Click(object sender, EventArgs e)
        //{
        //    // 1. Сохраняем листы
        //    AppSettings.SearchConfig.TargetSheetNames = _textBoxTargetSheets.Text
        //        .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
        //        .Select(s => s.Trim())
        //        .ToList();

        //    // <-- КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ СОХРАНЕНИЯ СЛОВ -->
        //    // 2. Сохраняем слова для поиска
        //    AppSettings.SearchConfig.SearchWords = _textBoxSearchWords.Text
        //        .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
        //        .Select(s => s.Trim())
        //        .ToList();
        //    // ----------------------------------------------

        //    // 3. Сохраняем все настройки на диск
        //    AppSettings.Save();

        //    UpdateStatus("Настройки успешно применены и сохранены.");
        //}
        //    // -----------------------------------------------

        //    UpdateStatus("Настройки успешно применены и сохранены.");
        //}

        // --- ОБНОВЛЕННЫЙ МЕТОД: Поиск и Подсчет (Сводная таблица по листам) ---
        // --- НОВЫЙ МЕТОД: Поиск и Вывод содержимого ячеек (Полный скан по листам) ---
        // --- ОБНОВЛЕННЫЙ МЕТОД: Поиск, подсчет и агрегация по найденным словам ---

        /// <summary>
        /// Считает количество вхождений подстроки в строку без учета регистра.
        /// </summary>
        private static int CountOccurrences(string source, string word)
        {
            if (string.IsNullOrEmpty(word) || string.IsNullOrEmpty(source))
            {
                return 0;
            }

            int count = 0;
            int index = -1;
            // Используем IndexOf с StringComparison.OrdinalIgnoreCase для поиска без учета регистра
            // и итеративно ищем все вхождения.
            while ((index = source.IndexOf(word, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                count++;
            }
            return count;
        }

        public class ExcelMatch
        {
            public string Sheet { get; set; } = ""; // <- Добавлено для устранения CS0649
            public string Value { get; set; } = ""; // <- Добавлено для устранения CS0649

        }

        // --- НОВЫЙ МЕТОД: Поиск, вывод содержимого и подсчет схожих ячеек ---
        // --- НОВЫЙ МЕТОД: Поиск, вывод содержимого и подсчет схожих ячеек ---
        // ИЗМЕНЕНИЕ 1: Меняем возвращаемый тип
        // =============================================================
        // МЕТОД 1: Загрузка "Карты Клемм" (Приоритетов) из спец. листа
        // =============================================================
        private void LoadPrioritiesFromVisioSheet(OfficeOpenXml.ExcelPackage package)
        {
            _visioPriorityMap.Clear();

            // 1. Ищем лист с приоритетами (по ключевому слову VISIO)
            var sheet = package.Workbook.Worksheets["1.ТЗ на ОБЪЕКТ ZONT для VISIO"];
            if (sheet == null)
            {
                sheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("VISIO"));
            }

            if (sheet == null) return; // Листа нет — приоритеты не работают, не страшно

            int startRow = 2; // Пропускаем шапку
            int endRow = sheet.Dimension?.End.Row ?? 100;

            for (int r = startRow; r <= endRow; r++)
            {
                // Колонка 1 (A): Наименование (напр. "Насос отопления")
                // Колонка 2 (B): Номер клеммы (напр. "3" или "Клемма 3")
                string name = sheet.Cells[r, 1].Text.Trim().ToLower();
                string valStr = sheet.Cells[r, 2].Text.Trim();

                if (string.IsNullOrEmpty(name)) continue;

                int priority = 9999;

                // Вытаскиваем только цифры из ячейки (чтобы "Кл. 3" превратилось в 3)
                string digits = new string(valStr.Where(char.IsDigit).ToArray());

                if (int.TryParse(digits, out int p))
                {
                    priority = p;
                }
                else
                {
                    // Если номера нет, берем номер строки * 100 (чтобы сохранить порядок таблицы)
                    priority = r * 100;
                }

                if (!_visioPriorityMap.ContainsKey(name))
                {
                    _visioPriorityMap.Add(name, priority);
                }
            }
        }

        // =============================================================
        // МЕТОД 2: Основной сканер (С внедренной логикой SchematicOrder)
        // =============================================================

        private bool IsRowExcluded(int currentRow, string exclusionRule)
        {
            if (string.IsNullOrWhiteSpace(exclusionRule)) return false;

            // Разбиваем строку по запятым
            var parts = exclusionRule.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                string p = part.Trim();

                // Проверяем формат "229-" (от числа и до конца)
                if (p.EndsWith("-"))
                {
                    string numberPart = p.TrimEnd('-');
                    if (int.TryParse(numberPart, out int startLimit))
                    {
                        if (currentRow >= startLimit) return true; // Исключить, если строка больше или равна
                    }
                }
                // Проверяем формат "10" (конкретная строка)
                else
                {
                    if (int.TryParse(p, out int specificRow))
                    {
                        if (currentRow == specificRow) return true; // Исключить конкретную строку
                    }
                }
            }
            return false;
        }

        private List<RawExcelHit> ScanSpecificSheet(string filePath)
        {
            var rawHits = new List<RawExcelHit>();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            try
            {
                // 1. Открываем Основной файл (где проверяем условия)
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    try { LoadPrioritiesFromVisioSheet(package); } catch { }

                    // 2. Открываем Дополнительный файл (откуда берем значения)
                    ExcelPackage auxPackage = null;
                    string auxPath = AppSettings.SearchConfig.AuxFilePath;

                    if (!string.IsNullOrEmpty(auxPath) && File.Exists(auxPath))
                    {
                        try
                        {
                            auxPackage = new ExcelPackage(new FileInfo(auxPath));
                        }
                        catch (Exception ex)
                        {
                            // Можно вывести сообщение, если файл битый
                            System.Diagnostics.Debug.WriteLine($"Ошибка открытия доп. файла: {ex.Message}");
                        }
                    }

                    try
                    {
                        var targetSheets = AppSettings.SearchConfig.TargetSheetNames;
                        if (targetSheets == null || !targetSheets.Any())
                        {
                            targetSheets = package.Workbook.Worksheets
                                .Where(w => !w.Name.ToUpper().Contains("VISIO"))
                                .Select(w => w.Name).ToList();
                        }

                        foreach (var sheetName in targetSheets)
                        {
                            // --- ОСНОВНОЙ ЛИСТ (Для условий) ---
                            var mainSheet = package.Workbook.Worksheets[sheetName];
                            if (mainSheet == null || mainSheet.Dimension == null) continue;

                            // --- ДОПОЛНИТЕЛЬНЫЙ ЛИСТ (Для данных) ---
                            ExcelWorksheet auxSheet = null;
                            if (auxPackage != null)
                            {
                                // Ищем лист без учета регистра (Sheet1 == sheet1)
                                auxSheet = auxPackage.Workbook.Worksheets
                                    .FirstOrDefault(w => w.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                            }

                            int startRow = mainSheet.Dimension.Start.Row;
                            int endRow = mainSheet.Dimension.End.Row;

                            // === ЦИКЛ ПО СТРОКАМ ===
                            for (int row = startRow; row <= endRow; row++)
                            {
                                foreach (var rule in AppSettings.SearchConfig.Rules)
                                {
                                    // Пропуск исключенных строк
                                    if (IsRowExcluded(row, rule.ExcludedRows)) continue;

                                    bool matchFound = false;
                                    string foundInMainTable = ""; // Для отладки или фоллбэка, если нужно

                                    // =========================================================
                                    // 1. ПРОВЕРКА УСЛОВИЯ (ТОЛЬКО В ОСНОВНОЙ ТАБЛИЦЕ)
                                    // =========================================================

                                    // Вариант А: Поиск по значению (например, "1" в колонке L)
                                    if (rule.SearchByValue)
                                    {
                                        int condColIndex = ExcelColumnLetterToNumber(rule.ConditionColumn);
                                        if (condColIndex > 0)
                                        {
                                            // Смотрим в MAIN SHEET
                                            string cellValue = mainSheet.Cells[row, condColIndex].Text?.Trim();
                                            string targetValue = string.IsNullOrEmpty(rule.ConditionValue) ? "1" : rule.ConditionValue;

                                            if (string.Equals(cellValue, targetValue, StringComparison.OrdinalIgnoreCase))
                                            {
                                                matchFound = true;
                                                // Запоминаем, что нашли в основной (на всякий случай), но пока не используем
                                                int nameCol = ExcelColumnLetterToNumber(rule.SearchColumn);
                                                if (nameCol > 0) foundInMainTable = mainSheet.Cells[row, nameCol].Text?.Trim();
                                            }
                                        }
                                    }
                                    // Вариант Б: Обычный поиск по слову (ключевые слова)
                                    else
                                    {
                                        // Сначала проверяем доп. условие (если есть галочка UseCondition)
                                        bool conditionPass = true;
                                        if (rule.UseCondition)
                                        {
                                            int condColIndex = ExcelColumnLetterToNumber(rule.ConditionColumn);
                                            if (condColIndex > 0)
                                            {
                                                // Смотрим в MAIN SHEET
                                                string actualValue = mainSheet.Cells[row, condColIndex].Text?.Trim();
                                                if (!string.Equals(actualValue, rule.ConditionValue, StringComparison.OrdinalIgnoreCase))
                                                    conditionPass = false;
                                            }
                                        }

                                        if (conditionPass)
                                        {
                                            string textToSearch = "";
                                            int colIndex = ExcelColumnLetterToNumber(rule.SearchColumn);

                                            // Читаем текст из MAIN SHEET для поиска ключевых слов
                                            if (!string.IsNullOrWhiteSpace(rule.SearchColumn) && colIndex > 0)
                                                textToSearch = mainSheet.Cells[row, colIndex].Text?.Trim();
                                            else
                                            {
                                                // Поиск по всей строке
                                                StringBuilder sb = new StringBuilder();
                                                for (int c = 1; c <= 20; c++) sb.Append(mainSheet.Cells[row, c].Text?.Trim() + " ");
                                                textToSearch = sb.ToString().Trim();
                                            }

                                            if (!string.IsNullOrEmpty(textToSearch) && !string.IsNullOrEmpty(rule.ExcelValue))
                                            {
                                                var keywords = rule.ExcelValue.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                                                foreach (var key in keywords)
                                                {
                                                    if (textToSearch.IndexOf(key.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        matchFound = true;
                                                        foundInMainTable = textToSearch; // Нашли совпадение
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // =========================================================
                                    // 2. ИЗВЛЕЧЕНИЕ РЕЗУЛЬТАТА (В ЗАВИСИМОСТИ ОТ НАСТРОЙКИ)
                                    // =========================================================
                                    if (matchFound)
                                    {
                                        string resultValue = "";

                                        // СЦЕНАРИЙ: Берем из ДОПОЛНИТЕЛЬНОЙ таблицы
                                        if (rule.ResultSource == DataSourceType.AuxFile)
                                        {
                                            if (auxSheet != null)
                                            {
                                                // Берем ТУ ЖЕ строку (row) и колонку поиска (SearchColumn)
                                                // (или TargetColumn, если логика подразумевает иное, но обычно берут из колонки "Где искать")
                                                int targetColIdx = ExcelColumnLetterToNumber(rule.SearchColumn);

                                                if (targetColIdx > 0)
                                                {
                                                    // Читаем из AUX SHEET
                                                    resultValue = auxSheet.Cells[row, targetColIdx].Text?.Trim();
                                                }
                                            }
                                            else
                                            {
                                                // Лист в доп файле не найден — результат ПУСТОЙ.
                                                // Не подменяем его данными из основной таблицы!
                                                // Можно добавить лог: Console.WriteLine($"Лист {sheetName} не найден в Aux");
                                            }
                                        }
                                        // СЦЕНАРИЙ: Берем из ОСНОВНОЙ таблицы (старый режим)
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(rule.VisioMasterName))
                                                resultValue = rule.VisioMasterName; // Жесткое имя
                                            else
                                            {
                                                // Берем из найденной ячейки основной таблицы
                                                // Если мы искали по значению "1", то берем текст из колонки SearchColumn
                                                if (rule.SearchByValue && string.IsNullOrEmpty(foundInMainTable))
                                                {
                                                    int nameCol = ExcelColumnLetterToNumber(rule.SearchColumn);
                                                    if (nameCol > 0) resultValue = mainSheet.Cells[row, nameCol].Text?.Trim();
                                                }
                                                else
                                                {
                                                    resultValue = foundInMainTable;
                                                }
                                            }
                                        }

                                        // Если результат пустой (например, в доп таблице в этой ячейке пусто) — пропускаем
                                        if (string.IsNullOrEmpty(resultValue)) continue;

                                        // --- ДАЛЕЕ СТАНДАРТНАЯ ОБРАБОТКА (Приоритеты и добавление) ---
                                        int finalPriority = 9999;
                                        string nameLower = resultValue.ToLower().Trim();

                                        if (_visioPriorityMap.ContainsKey(nameLower))
                                        {
                                            finalPriority = _visioPriorityMap[nameLower];
                                        }
                                        else
                                        {
                                            var match = _visioPriorityMap.Keys
                                                .Where(k => nameLower.Contains(k))
                                                .OrderByDescending(k => k.Length)
                                                .FirstOrDefault();
                                            if (match != null) finalPriority = _visioPriorityMap[match];
                                        }

                                        rawHits.Add(new RawExcelHit
                                        {
                                            SheetName = sheetName,
                                            RowIndex = row,
                                            FullItemName = resultValue,       // <-- Значение из нужной таблицы
                                            TargetMasterName = resultValue,   // <-- Значение из нужной таблицы
                                            SearchTerm = rule.ExcelValue ?? "ConditionMatch",
                                            ConditionMet = true,
                                            Quantity = 1,
                                            IsLimited = rule.LimitQuantity,
                                            FoundRule = rule,
                                            SchematicOrder = finalPriority,
                                            TargetColumn = rule.TargetColumn
                                        });
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        // Закрываем доп. файл, чтобы не висел в памяти
                        if (auxPackage != null) auxPackage.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сканирования: {ex.Message}");
            }

            return rawHits;
        }

        // Вспомогательный метод для парсинга количества (добавьте его в класс Form1)
        private bool TryParseQuantity(string text, out int result)
        {
            result = 1;
            if (string.IsNullOrWhiteSpace(text)) return false;

            // Заменяем точку на запятую для корректного парсинга в RU-локали
            text = text.Replace(".", ",");

            if (double.TryParse(text, System.Globalization.NumberStyles.Any, new System.Globalization.CultureInfo("ru-RU"), out double d))
            {
                if (d > 0)
                {
                    result = (int)Math.Round(d); // Округляем (на случай 1.00)
                    return true;
                }
            }
            return false;
        }
        // ПРИМЕЧАНИЕ: Предполагается, что функция GetColumnIndex(string columnName) 
        // и структура RawExcelHit уже определены.

        // Удалите или проигнорируйте вспомогательный метод CountOccurrences, он больше не нужен
        // private static int CountOccurrences(...) { ... }
        // ----------------------------------------------------------------------


        private void UpdateDataGridView()
        {
            dataGridView.Rows.Clear();
            foreach (var rowData in data)
            {
                var row = new DataGridViewRow();
                foreach (var colName in COLS_OUT)
                {
                    row.Cells.Add(new DataGridViewTextBoxCell { Value = rowData.GetValueOrDefault(colName, string.Empty) });
                }
                dataGridView.Rows.Add(row);
            }
            dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private string GetDeclension(int number, string nominative, string genitiveSingular, string genitivePlural)
        {
            number = Math.Abs(number) % 100;
            var n1 = number % 10;
            if (number > 10 && number < 20) return genitivePlural;
            if (n1 > 1 && n1 < 5) return genitiveSingular;
            if (n1 == 1) return nominative;
            return genitivePlural;
        }

        private void ShowResultMessage(int totalFound)
        {

            int count = _rawHits.Count;
            string word = GetDeclension(count, "позиция", "позиции", "позиций");
            if (count > 0)
            {
                UpdateStatus($"? Сканирование завершено. Найдено {count} {word}.");
            }
            else
            {
                UpdateStatus("? Сканирование завершено. Ничего не найдено.");
                MessageBox.Show("Ничего не найдено. Проверьте настройки поиска.", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }



        private void GeneratePage(Visio.Document doc, string pageName, Dictionary<string, int> masterCounts, List<SearchRule> SearchRules)
        {
            // ИСПРАВЛЕНО: используем переданный аргумент SearchRules вместо несуществующего config
            var masterMap = RulesToMap(SearchRules);

            // 1. Создание нового листа (если его нет) или выбор активного
            Visio.Page page;
            try
            {
                page = doc.Pages.Add();
                page.Name = pageName;
            }
            catch (COMException)
            {
                // Если лист уже есть, используем его
                page = doc.Pages.get_ItemU(pageName);
                page.Background = 0;
            }

            // Параметры размещения
            double currentX = 0.5; // Начальная позиция X (в метрах)
            double currentY = 0.5; // Начальная позиция Y (в метрах)
            double rowHeight = 0; // Высота самого высокого элемента в текущей строке

            // ИСПРАВЛЕНО: Закомментированы неиспользуемые переменные (CS0219)
            // const double SPACING = 0.05; 
            // const double PAGE_WIDTH = 0.279; 
            const double SPACING = 0.05; // Если вы планируете использовать их, раскомментируйте и добавьте логику
            const double PAGE_WIDTH = 0.279;

            // Проходим по всем элементам, которые нужно добавить (из Excel)
            foreach (var excelKeyCount in masterCounts)
            {
                string excelKey = excelKeyCount.Key;
                int count = excelKeyCount.Value;

                // 2. Находим имя мастера Visio по ключу из Excel
                if (masterMap.TryGetValue(excelKey, out string masterName))
                {
                    // Получаем ссылку на мастер из открытых трафаретов
                    Visio.Master master = doc.Masters.get_ItemU(masterName);

                    // Добавляем фигуры
                    for (int i = 0; i < count; i++)
                    {
                        // Добавляем фигуру на страницу
                        Visio.Shape shape = page.Drop(master, 0, 0); // Изначально бросаем в (0,0)

                        // Получаем текущие размеры фигуры (ширина/высота)
                        double shapeWidth = shape.CellsU["Width"].ResultIU;
                        double shapeHeight = shape.CellsU["Height"].ResultIU;

                        // Проверяем, помещается ли фигура в текущей строке
                        if (currentX + shapeWidth > PAGE_WIDTH)
                        {
                            // Переход на новую строку
                            currentX = 0.5;
                            currentY += rowHeight + SPACING; // Сдвиг вниз
                            rowHeight = 0; // Сбрасываем высоту
                        }

                        // Устанавливаем позицию фигуры (центр)
                        shape.CellsU["PinX"].ResultIU = currentX + shapeWidth / 2.0;
                        shape.CellsU["PinY"].ResultIU = currentY + shapeHeight / 2.0;

                        // Обновляем параметры для следующей фигуры
                        currentX += shapeWidth + SPACING;
                        if (shapeHeight > rowHeight)
                            rowHeight = shapeHeight;
                    }
                }
            }
            // Опционально: подгоняем размер страницы под содержимое
            page.ResizeToFitContents();
        }

        private void ShowResult(bool success, string message)
        {
            // Проверка, нужен ли Invoke (если метод вызван не из UI-потока)
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ShowResult(success, message)));
                return;
            }

            // Вывод сообщения
            if (success)
            {
                MessageBox.Show("Документ Visio успешно создан и открыт.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"Произошла ошибка при генерации документа Visio: {message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Включаем кнопки обратно, если вы их отключали
            // btnOpenVisio.Enabled = true;
            // btnLoad.Enabled = true;
        }



        public static class VisioVbaRunner
        {
            // *** ВАЖНО: Этот метод освобождает COM-объекты. Не вызывайте его для doc или app, пока они нужны! ***
            private static void ReleaseComObject(object? obj)
            {
                try { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); } catch { }
            }

            public static void RunDrawingMacro(Visio.Document doc, Visio.Page page, List<VisioItem> itemsToDraw, VisioConfiguration config)
            {
                if (itemsToDraw == null || itemsToDraw.Count == 0) return;

                string pageName = page.Name;
                string moduleName = "Mod_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                StringBuilder sb = new StringBuilder();

                const double MM_TO_INCH = 0.0393701;

                // --- НАСТРОЙКИ SEQUENTIAL ---
                var seq = config.SequentialDrawing;

                // Парсим стартовые координаты
                double seqStartX = 10.0, seqStartY = 250.0;
                var parts = seq.StartCoordinatesXY.Split(',');
                if (parts.Length >= 2)
                {
                    double.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out seqStartX);
                    double.TryParse(parts[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out seqStartY);
                }

                // --- ГЕНЕРАЦИЯ VBA ---
                sb.AppendLine($"Sub Draw_{moduleName}()");
                sb.AppendLine("    On Error GoTo HandleError"); // Добавлено для явной диагностики VBA ошибок
                sb.AppendLine("    Const visInches = 65");
                sb.AppendLine($"    Dim pg As Visio.Page");
                sb.AppendLine($"    Dim doc As Visio.Document");
                sb.AppendLine($"    Dim mst As Visio.Master");
                sb.AppendLine($"    Dim shp As Visio.Shape");
                sb.AppendLine($"    Dim i As Integer");
                sb.AppendLine($"    Dim mName As String"); // Добавлено для поиска мастера
                sb.AppendLine($"    Dim dropX As Double, dropY As Double");
                sb.AppendLine($"    Dim w As Double, h As Double");

                sb.AppendLine($"    Set pg = ActiveDocument.Pages.ItemU(\"{pageName}\")");
                sb.AppendLine($"    If pg Is Nothing Then GoTo CleanExit ' Проверка страницы");

                // Массивы данных (из вашего кода)
                int count = itemsToDraw.Count;
                sb.AppendLine($"    Dim masters({count}) As String");
                sb.AppendLine($"    Dim types({count}) As String");
                sb.AppendLine($"    Dim xPos({count}) As Double");
                sb.AppendLine($"    Dim yPos({count}) As Double");
                sb.AppendLine($"    Dim anchors({count}) As String");

                // ... заполнение массивов ...
                for (int j = 0; j < count; j++) // Изменена i на j, чтобы не конфликтовать с i в VBA цикле
                {
                    var item = itemsToDraw[j];
                    // Использовать замену кавычек как в вашем рабочем коде
                    string safeMasterName = item.MasterName.Replace("\"", "\"\"");
                    safeMasterName = safeMasterName.Trim();
                    sb.AppendLine($"    masters({j}) = \"{safeMasterName}\"");
                    sb.AppendLine($"    types({j}) = \"{item.PlacementType}\"");
                    sb.AppendLine($"    anchors({j}) = \"{item.Anchor}\"");

                    if (item.PlacementType == "Manual")
                    {
                        sb.AppendLine($"    xPos({j}) = {(item.X * MM_TO_INCH).ToString(CultureInfo.InvariantCulture)}");
                        sb.AppendLine($"    yPos({j}) = {(item.Y * MM_TO_INCH).ToString(CultureInfo.InvariantCulture)}");
                    }
                }

                sb.AppendLine("");
                sb.AppendLine("    Const MM2IN = 0.0393701");

                // Переменные курсора (Sequential)
                sb.AppendLine($"    Dim curX As Double: curX = {seqStartX.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim curY As Double: curY = {seqStartY.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim lineStart As Double: lineStart = curX");
                sb.AppendLine($"    Dim maxW As Double: maxW = {seq.MaxLineWidthMM.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim hGap As Double: hGap = {seq.HorizontalStepMM.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim vGap As Double: vGap = {seq.VerticalStepMM.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine("    Dim rowMaxH As Double: rowMaxH = 0");

                sb.AppendLine($"    For i = 0 To {count - 1}");

                // --- НАДЕЖНАЯ ЛОГИКА ПОИСКА МАСТЕРА --- (в документе, затем в трафаретах)
                sb.AppendLine("        Set mst = Nothing");
                sb.AppendLine("        mName = masters(i)");

                // 1. Поиск в документе (по NameU)
                sb.AppendLine("        On Error Resume Next");
                sb.AppendLine("        Set mst = ActiveDocument.Masters.ItemU(mName)");
                sb.AppendLine("        On Error GoTo 0");

                // 2. Если не нашли, ищем во ВСЕХ открытых трафаретах
                sb.AppendLine("        If mst Is Nothing Then");
                sb.AppendLine("            For Each doc In Application.Documents");
                sb.AppendLine("                If doc.Type = 2 Then ' 2 = visTypeStencil");
                sb.AppendLine("                    On Error Resume Next");
                sb.AppendLine("                    Set mst = doc.Masters.ItemU(mName)");
                sb.AppendLine("                    On Error GoTo 0");
                sb.AppendLine("                    If Not mst Is Nothing Then Exit For");
                sb.AppendLine("                End If");
                sb.AppendLine("            Next doc");
                sb.AppendLine("        End If");

                sb.AppendLine("        If mst Is Nothing Then");
                sb.AppendLine("            Debug.Print \"[Drawing Skip] Master NOT FOUND: \" & mName");
                sb.AppendLine("        Else ' Если мастер найден, выполняем логику размещения");
                sb.AppendLine("            Dim w As Double, h As Double");
                sb.AppendLine("            w = mst.Cells(\"Width\").Result(visInches)");
                sb.AppendLine("            h = mst.Cells(\"Height\").Result(visInches)");
                sb.AppendLine("            Dim dropX As Double, dropY As Double");

                // --- ЛОГИКА КООРДИНАТ (Sequential/Manual) --- (оставлена без изменений)
                sb.AppendLine("            If types(i) = \"Sequential\" Then");
                sb.AppendLine("                If (curX + w - lineStart) > maxW Then");
                sb.AppendLine("                    curX = lineStart");
                sb.AppendLine("                    curY = curY - rowMaxH - vGap");
                sb.AppendLine("                    rowMaxH = 0");
                sb.AppendLine("                End If");
                sb.AppendLine("                dropX = curX + (w / 2)");
                sb.AppendLine("                dropY = curY - (h / 2)");
                sb.AppendLine("                curX = curX + w + hGap");
                sb.AppendLine("                If h > rowMaxH Then rowMaxH = h");
                sb.AppendLine("            Else ' Manual");
                sb.AppendLine("                dropX = xPos(i)");
                sb.AppendLine("                dropY = yPos(i)");
                sb.AppendLine("                Select Case anchors(i)");
                sb.AppendLine("                    Case \"TopLeft\": dropX = dropX + (w / 2): dropY = dropY - (h / 2)");
                sb.AppendLine("                    Case \"BottomLeft\": dropX = dropX + (w / 2): dropY = dropY + (h / 2)");
                sb.AppendLine("                End Select");
                sb.AppendLine("            End If");

                // ВСТАВКА
                sb.AppendLine("            Set shp = pg.Drop(mst, dropX, dropY)");

                // --- Фиксация Pin (Удален некорректный блок смещения PinY) ---
                sb.AppendLine("            If Not shp Is Nothing Then");
                sb.AppendLine("                shp.Cells(\"PinX\").ResultIU = dropX");
                sb.AppendLine("                shp.Cells(\"PinY\").ResultIU = dropY");
                sb.AppendLine("            End If");

                sb.AppendLine("        End If"); // Конец If Not mst Is Nothing Then
                sb.AppendLine("    Next i");

                sb.AppendLine("CleanExit:"); // Точка выхода при отсутствии ошибок
                sb.AppendLine("    Set pg = Nothing: Set mst = Nothing: Set shp = Nothing: Set doc = Nothing"); // Очистка COM-переменных
                sb.AppendLine("    Exit Sub");

                sb.AppendLine("HandleError:"); // Обработчик ошибок
                sb.AppendLine("    MsgBox \"VBA Run-time error: \" & Err.Description & \" (Code: \" & Err.Number & \") on line \" & Erl, vbCritical");
                sb.AppendLine("    Resume CleanExit");

                sb.AppendLine("End Sub");

                // Запуск
                Microsoft.Vbe.Interop.VBComponent? vbComp = null;
                try
                {
                    if (doc.VBProject != null)
                    {
                        vbComp = doc.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                        vbComp.Name = moduleName;
                        vbComp.CodeModule.AddFromString(sb.ToString());
                        // ВАЖНО: ExecuteLine может выбрасывать исключение, если есть проблема с COM (Invalid DOS Handle)
                        doc.ExecuteLine($"Draw_{moduleName}");
                    }
                }
                catch (Exception ex)
                {
                    // Здесь будет поймана ошибка "Недопустимый дескриптор DOS"
                    MessageBox.Show("КРИТИЧЕСКАЯ ОШИБКА COM INTEROP: " + ex.Message +
                        "\n\nВозможно, объект Visio.Application или Visio.Document был освобожден в ВЫЗЫВАЮЩЕМ коде (через Marshal.ReleaseComObject или завершение using-блока) до того, как макрос завершил работу.",
                        "Ошибка VBA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // !!! КРИТИЧЕСКИ ВАЖНО ДЛЯ ОЧИСТКИ !!!
                    if (doc.VBProject != null && vbComp != null)
                    {
                        try { doc.VBProject.VBComponents.Remove(vbComp); } catch { }
                    }
                    ReleaseComObject(vbComp);
                }
            }
        }

        // Вспомогательный класс для передачи данных
        public class VisioItem
        {
            public string MasterName { get; set; }
            public double X { get; set; } = 0; // Координаты (в MM, конвертируются в дюймы в VBA Runner)
            public double Y { get; set; } = 0; // Координаты (в MM, конвертируются в дюймы в VBA Runner)
            public string Anchor { get; set; } = "Center";

            public string PlacementType { get; set; } = "Manual";
        }

        private async void OpenVisioClick(object? sender, EventArgs e)
        {
            // Проверяем, создана ли таблица
            if (mainDgv == null || mainDgv.Rows.Count == 0)
            {
                MessageBox.Show("Сначала создайте таблицу (кнопка 'Создать таблицу')!", "Ошибка");
                return;
            }

            this.Enabled = false;
            UpdateStatus("⏳ Подготовка данных из таблицы...");

            var tableHits = new List<RawExcelHit>();

            foreach (DataGridViewRow row in mainDgv.Rows)
            {
                if (row.IsNewRow) continue;

                // 1. Читаем имя блока (Колонка 3 - "BlockName" или индекс 2)
                var cellValue = row.Cells["BlockName"].Value?.ToString(); // Или row.Cells[2].Value

                // 2. Читаем номер п/п для сортировки (Колонка 1 - "№" или индекс 0)
                // Если там пусто или не число, ставим 0 или int.MaxValue (в конец)
                int sortOrder = int.MaxValue;
                var sortVal = row.Cells[0].Value?.ToString(); // Предполагаем, что № в 0-й колонке
                if (int.TryParse(sortVal, out int parsedOrder))
                {
                    sortOrder = parsedOrder;
                }

                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    tableHits.Add(new RawExcelHit
                    {
                        SearchTerm = cellValue.Trim(),
                        Quantity = 1,
                        ConditionMet = true,
                        SortIndex = sortOrder // Запоминаем номер
                    });
                }
            }

            // ВАЖНО: Сортируем список по возрастанию номера п/п перед отправкой
            tableHits = tableHits.OrderBy(h => h.SortIndex).ToList();

            if (!tableHits.Any())
            {
                MessageBox.Show("В таблице нет данных в колонке 'Наименование блока'.");
                this.Enabled = true;
                return;
            }

            UpdateStatus("⏳ Запуск Visio...");

            await Task.Run(() =>
            {
                Visio.Application? visioApp = null;
                try
                {
                    visioApp = new Visio.Application();
                    visioApp.Visible = true;
                    var doc = visioApp.Documents.Add("");

                    // Открываем трафареты из всех конфигов
                    var allStencils = AppSettings.LabelingConfig.StencilFilePaths
                        .Union(AppSettings.SchemeConfig.StencilFilePaths)
                        .Union(AppSettings.CabinetConfig.StencilFilePaths)
                        .Distinct()
                        .Where(File.Exists)
                        .ToList();

                    foreach (var path in allStencils)
                    {
                        visioApp.Documents.OpenEx(path, (short)Visio.VisOpenSaveArgs.visOpenDocked);
                    }

                    // === ГЕНЕРАЦИЯ 3-Х ЛИСТОВ ===
                    // Функция GeneratePageDirectly сама проверит:
                    // 1. Есть ли имя блока из таблицы в маппинге этого конфига (Tab 2)
                    // 2. Если есть -> возьмет MasterName
                    // 3. Возьмет настройки координат/отступов из SequentialDrawing (Tab 3) и разместит

                    UpdateStatus("Рисуем Маркировку...");
                    GeneratePageDirectly(doc, "Маркировка", tableHits, AppSettings.LabelingConfig);

                    UpdateStatus("Рисуем Схему...");
                    GeneratePageDirectly(doc, "Схема", tableHits, AppSettings.SchemeConfig);

                    UpdateStatus("Рисуем Шкаф...");
                    GeneratePageDirectly(doc, "Шкаф", tableHits, AppSettings.CabinetConfig);

                    UpdateStatus("✅ Visio готово.");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"❌ Ошибка Visio: {ex.Message}");
                    MessageBox.Show(ex.Message);
                }
            });

            this.Enabled = true;
        }

        // Внутри класса Form1

        private void GeneratePageDirectly(Visio.Document doc, string pageName, List<RawExcelHit> hits, VisioConfiguration config)
        {
            Visio.Page page = null;
            try
            {
                try { page = doc.Pages.get_ItemU(pageName); }
                catch { page = doc.Pages.Add(); page.Name = pageName; }
                SetupVisioPage(page, config);
            }
            catch { return; }

            // --- 1. ФИКСИРОВАННЫЕ ФИГУРЫ (Predefined) ---
            if (config.PredefinedMasterConfigs != null)
            {
                foreach (var fixedItem in config.PredefinedMasterConfigs)
                {
                    if (string.IsNullOrWhiteSpace(fixedItem.MasterName)) continue;

                    double x = 0, y = 0;
                    var coords = fixedItem.CoordinatesXY?.Split(new[] { ',', ';' });
                    if (coords != null && coords.Length >= 2)
                    {
                        double.TryParse(coords[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out x);
                        double.TryParse(coords[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out y);
                    }

                    for (int i = 0; i < fixedItem.Quantity; i++)
                    {
                        Visio.Shape shp = DropShapeOnPage(page, fixedItem.MasterName, 0, 0);

                        if (shp != null)
                        {
                            string anchor = !string.IsNullOrWhiteSpace(fixedItem.Anchor) ? fixedItem.Anchor : "Center";
                            SetShapePosition(shp, x, y, anchor);
                        }
                    }
                }
            }

            // --- 2. НАЙДЕННЫЕ ФИГУРЫ (Поиск по карте) ---
            // --- 2. НАЙДЕННЫЕ ФИГУРЫ (Sequential) ---
            if (config.SearchRules != null && config.SearchRules.Any() && hits != null && hits.Any())
            {
                bool seqEnabled = config.SequentialDrawing.Enabled;

                // Начальные координаты в ММ
                double curX_MM = 10;
                double curY_MM = 200;

                // Парсим стартовые координаты из настроек
                var sCoords = config.SequentialDrawing.StartCoordinatesXY?.Split(new[] { ',', ';' });
                if (sCoords != null && sCoords.Length >= 2)
                {
                    double.TryParse(sCoords[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out curX_MM);
                    double.TryParse(sCoords[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out curY_MM);
                }

                double maxW_MM = config.SequentialDrawing.MaxLineWidthMM;
                double hGap_MM = config.SequentialDrawing.HorizontalStepMM;
                double vGap_MM = config.SequentialDrawing.VerticalStepMM;

                // Глобальный якорь из выпадающего списка (например "TopLeft" или "Center")
                string globalAnchor = !string.IsNullOrWhiteSpace(config.SequentialDrawing.Anchor)
                                      ? config.SequentialDrawing.Anchor
                                      : "Center";

                double startX_MM = curX_MM;
                double rowMaxH_MM = 0;
                int qfCounter = 2;         // Автоматы (QF2...)
                int blockLabelCounter = 3; // Клеммы (3, 4...)
                int sensorCounter = 1;     // Датчики (1, 2, 3...)

                // Ключевые слова для разделения нумерации клемм
                string[] splitKeywords = new[] { "Перекл.", "Смеситель", "Управл.", "давления" };

                foreach (var hit in hits)
                {
                    if (!hit.ConditionMet || string.IsNullOrWhiteSpace(hit.SearchTerm)) continue;

                    var matchedRule = config.SearchRules.FirstOrDefault(r =>
                        !string.IsNullOrEmpty(r.ExcelValue) &&
                        hit.SearchTerm.IndexOf(r.ExcelValue, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (matchedRule != null && !string.IsNullOrWhiteSpace(matchedRule.VisioMasterName))
                    {
                        int countToDraw = matchedRule.LimitQuantity ? 1 : hit.Quantity;

                        for (int i = 0; i < countToDraw; i++)
                        {
                            // 1. Бросаем ОСНОВНУЮ фигуру
                            Visio.Shape shp = DropShapeOnPage(page, matchedRule.VisioMasterName, 0, 0);

                            if (shp != null)
                            {
                                // --- А. ЗАПОЛНЕНИЕ ДАННЫХ ---

                                // 1. Наименование
                                SetShapeData(shp, "Наименование блоков", hit.SearchTerm);

                                // 2. Нумерация автоматов (если есть поле "Номер автомата")
                                if (SetShapeData(shp, "Номер автомата", $"QF{qfCounter}"))
                                {
                                    qfCounter++;
                                }

                                // 3. === НОВОЕ: Нумерация датчиков (если есть поле "Номер датчика") ===
                                // Начинаем с 1. Если записали успешно — увеличиваем.
                                if (SetShapeData(shp, "Номер датчика", sensorCounter.ToString()))
                                {
                                    sensorCounter++;
                                }

                                // --- Б. ПОЗИЦИОНИРОВАНИЕ ОСНОВНОЙ ФИГУРЫ ---
                                string finalAnchor = !string.IsNullOrWhiteSpace(matchedRule.Anchor) ? matchedRule.Anchor : globalAnchor;

                                double mainW_MM = 0;
                                double mainH_MM = 0;
                                double mainPinX_MM = 0;
                                double mainPinY_MM = 0;

                                if (seqEnabled)
                                {
                                    // Змейка
                                    mainW_MM = shp.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters];
                                    mainH_MM = shp.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];
                                    double rightBoundary = startX_MM + maxW_MM;

                                    if (curX_MM + mainW_MM > rightBoundary)
                                    {
                                        curX_MM = startX_MM;
                                        double stepDown = (rowMaxH_MM > 0 ? rowMaxH_MM : mainH_MM) + vGap_MM;
                                        curY_MM -= stepDown;
                                        rowMaxH_MM = 0;
                                    }

                                    mainPinX_MM = curX_MM + (mainW_MM / 2.0);
                                    mainPinY_MM = curY_MM - (mainH_MM / 2.0);

                                    shp.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = mainPinX_MM;
                                    shp.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = mainPinY_MM;

                                    curX_MM += mainW_MM + hGap_MM;
                                    if (mainH_MM > rowMaxH_MM) rowMaxH_MM = mainH_MM;
                                }
                                else
                                {
                                    // Статика
                                    double rX = 0, rY = 0;
                                    var rCoords = matchedRule.CoordinatesXY?.Split(',');
                                    if (rCoords != null && rCoords.Length >= 2)
                                    {
                                        double.TryParse(rCoords[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out rX);
                                        double.TryParse(rCoords[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out rY);
                                    }
                                    SetShapePosition(shp, rX, rY, finalAnchor);

                                    mainW_MM = shp.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters];
                                    mainH_MM = shp.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];
                                    mainPinX_MM = shp.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters];
                                    mainPinY_MM = shp.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters];
                                }

                                // --- В. ДОБАВЛЕНИЕ НУМЕРАЦИИ КЛЕММ (Только для листа "Схемы") ---
                                if (pageName.IndexOf("Схем", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // Проверяем, нужно ли делить блок на 2 части
                                    bool splitBlock = splitKeywords.Any(k => hit.SearchTerm.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);

                                    if (splitBlock)
                                    {
                                        // === ВАРИАНТ А: ДВЕ ПОЛОВИНКИ ===
                                        double halfWidth = mainW_MM / 2.0;
                                        double topEdgeY = mainPinY_MM + (mainH_MM / 2.0);

                                        // Левая половинка
                                        Visio.Shape leftLbl = DropShapeOnPage(page, "0.Нумера БЛОКОВ таблицы", 0, 0);
                                        if (leftLbl != null)
                                        {
                                            leftLbl.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters] = halfWidth;
                                            double lblH = leftLbl.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                            double leftPinX = mainPinX_MM - (mainW_MM / 4.0);
                                            double lblPinY = topEdgeY - (lblH / 2.0);

                                            leftLbl.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = leftPinX;
                                            leftLbl.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = lblPinY;

                                            SetShapeData(leftLbl, "Нумерация клемм", blockLabelCounter.ToString());
                                            blockLabelCounter++;
                                        }

                                        // Правая половинка
                                        Visio.Shape rightLbl = DropShapeOnPage(page, "0.Нумера БЛОКОВ таблицы", 0, 0);
                                        if (rightLbl != null)
                                        {
                                            rightLbl.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters] = halfWidth;
                                            double lblH = rightLbl.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                            double rightPinX = mainPinX_MM + (mainW_MM / 4.0);
                                            double lblPinY = topEdgeY - (lblH / 2.0);

                                            rightLbl.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = rightPinX;
                                            rightLbl.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = lblPinY;

                                            SetShapeData(rightLbl, "Нумерация клемм", blockLabelCounter.ToString());
                                            blockLabelCounter++;
                                        }
                                    }
                                    else
                                    {
                                        // === ВАРИАНТ Б: ОДИН ЦЕЛЫЙ БЛОК ===
                                        Visio.Shape labelShp = DropShapeOnPage(page, "0.Нумера БЛОКОВ таблицы", 0, 0);
                                        if (labelShp != null)
                                        {
                                            labelShp.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters] = mainW_MM;
                                            double labelH_MM = labelShp.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                            double topEdgeY = mainPinY_MM + (mainH_MM / 2.0);
                                            double newLabelPinY = topEdgeY - (labelH_MM / 2.0);

                                            labelShp.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = mainPinX_MM;
                                            labelShp.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = newLabelPinY;

                                            SetShapeData(labelShp, "Нумерация клемм", blockLabelCounter.ToString());
                                            blockLabelCounter++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

                // Удаление пустой первой страницы
                try
            {
                // Проверяем первый лист (индексация в Visio с 1)
                if (doc.Pages.Count > 1)
                {
                    Visio.Page firstPage = doc.Pages[1];

                    // Удаляем ТОЛЬКО если имя похоже на стандартное ("Page-1", "Страница-1") И она пустая
                    // Это защитит ваши листы "Маркировка" и т.д. от удаления
                    string name = firstPage.Name;
                    if ((name.StartsWith("Page") || name.StartsWith("Страница")) && firstPage.Shapes.Count == 0)
                    {
                        firstPage.Delete(0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при удалении пустой страницы: " + ex.Message);
            }
            // ======================================================

            UpdateStatus("✅ Visio готово.");
        }

        // --- УЛУЧШЕННЫЙ МЕТОД ВСТАВКИ ФИГУРЫ ---
        private Visio.Shape? DropShapeOnPage(Visio.Page page, string masterName, double xMM, double yMM)
        {
            // 1. Проверка входных данных
            if (string.IsNullOrWhiteSpace(masterName)) return null;
            masterName = masterName.Trim();

            Visio.Master? mst = null;
            Visio.Document doc = page.Document;
            Visio.Application app = doc.Application;

            // --- ЛОКАЛЬНАЯ ФУНКЦИЯ ДЛЯ ПОИСКА ---
            // Пытается найти сначала по универсальному имени (NameU), затем по локальному (Name)
            Visio.Master? TryGetMaster(Visio.Document d, string name)
            {
                // 1. Приоритет: Универсальное имя (ItemU)
                try
                {
                    return d.Masters.get_ItemU(name);
                }
                catch { }

                // 2. Резерв: Локальное имя через индексатор [ ]
                // Важно: здесь нельзя использовать get_Item()
                try
                {
                    return d.Masters[name];
                }
                catch { }

                return null;
            }
            // -------------------------------------

            // 2. Ищем в текущем документе
            mst = TryGetMaster(doc, masterName);

            // 3. Если не нашли, перебираем открытые трафареты
            if (mst == null)
            {
                foreach (Visio.Document d in app.Documents)
                {
                    // Работаем только с трафаретами (.vssx, .vss)
                    if (d.Type == Visio.VisDocumentTypes.visTypeStencil)
                    {
                        mst = TryGetMaster(d, masterName);
                        if (mst != null) break; // Нашли — выходим из цикла
                    }
                }
            }

            // 4. Если мастер так и не найден — выход
            if (mst == null) return null;

            // 5. Вставка фигуры (конвертация MM -> Inch)
            try
            {
                const double MM_TO_INCH = 1.0 / 25.4;
                return page.Drop(mst, xMM * MM_TO_INCH, yMM * MM_TO_INCH);
            }
            catch (Exception)
            {
                // Логирование ошибки можно добавить сюда
                return null;
            }
        }

        private void SetShapePosition(Visio.Shape shp, double xMM, double yMM, string anchor)
{
    if (shp == null) return;

    // 1. Формируем строки с единицами измерения "mm" для Visio
    // Это самый надежный способ избежать путаницы дюймы/мм
    string sX = xMM.ToString(CultureInfo.InvariantCulture) + " mm";
    string sY = yMM.ToString(CultureInfo.InvariantCulture) + " mm";

    // 2. Настраиваем Anchor (Точку привязки ВНУТРИ фигуры - LocPin)
    // FormulaU позволяет писать "Width*0" и т.д.
    switch (anchor?.ToLower())
    {
        case "topleft":
            shp.CellsU["LocPinX"].FormulaU = "Width*0"; // Левый край
            shp.CellsU["LocPinY"].FormulaU = "Height*1"; // Верхний край
            break;
        case "topcenter":
            shp.CellsU["LocPinX"].FormulaU = "Width*0.5";
            shp.CellsU["LocPinY"].FormulaU = "Height*1";
            break;
        case "topright":
            shp.CellsU["LocPinX"].FormulaU = "Width*1";
            shp.CellsU["LocPinY"].FormulaU = "Height*1";
            break;
        case "centerleft":
            shp.CellsU["LocPinX"].FormulaU = "Width*0";
            shp.CellsU["LocPinY"].FormulaU = "Height*0.5";
            break;
        case "centerright":
            shp.CellsU["LocPinX"].FormulaU = "Width*1";
            shp.CellsU["LocPinY"].FormulaU = "Height*0.5";
            break;
        case "bottomleft":
            shp.CellsU["LocPinX"].FormulaU = "Width*0";
            shp.CellsU["LocPinY"].FormulaU = "Height*0";
            break;
        case "bottomcenter":
            shp.CellsU["LocPinX"].FormulaU = "Width*0.5";
            shp.CellsU["LocPinY"].FormulaU = "Height*0";
            break;
        case "bottomright":
            shp.CellsU["LocPinX"].FormulaU = "Width*1";
            shp.CellsU["LocPinY"].FormulaU = "Height*0";
            break;
        case "center":
        default:
            shp.CellsU["LocPinX"].FormulaU = "Width*0.5";
            shp.CellsU["LocPinY"].FormulaU = "Height*0.5";
            break;
    }

    // 3. Ставим саму фигуру в координаты на листе (Pin)
    shp.CellsU["PinX"].FormulaU = sX;
    shp.CellsU["PinY"].FormulaU = sY;
}

        // Вспомогательный метод для вставки одной фигуры
        private Visio.Shape? DropShapeOnPage(Visio.Page page, string masterName, double xMM, double yMM, int qty, bool isSequential = false)
        {
            if (string.IsNullOrWhiteSpace(masterName)) return null;

            masterName = masterName.Trim();
            Visio.Master? mst = null;
            Visio.Document doc = page.Document;
            Visio.Application app = doc.Application;

            // --- 1. Ищем мастер в самом документе ---

            // Попытка А: По универсальному имени (NameU)
            try { mst = doc.Masters.get_ItemU(masterName); } catch { }

            // Попытка Б: По локальному имени (Name) через индексатор [ ]
            if (mst == null)
            {
                try { mst = doc.Masters[masterName]; } catch { }
            }

            // --- 2. Если нет, ищем во всех открытых трафаретах ---
            if (mst == null)
            {
                foreach (Visio.Document d in app.Documents)
                {
                    // Проверяем только трафареты (Type = 2, visTypeStencil)
                    if (d.Type == Visio.VisDocumentTypes.visTypeStencil)
                    {
                        // Попытка А: NameU
                        try { mst = d.Masters.get_ItemU(masterName); } catch { }

                        // Попытка Б: Name через индексатор [ ]
                        if (mst == null)
                        {
                            try { mst = d.Masters[masterName]; } catch { }
                        }

                        if (mst != null) break; // Нашли!
                    }
                }
            }

            // --- 3. Если так и не нашли ---
            if (mst == null)
            {
                // Можно раскомментировать для отладки
                // Console.WriteLine($"Мастер '{masterName}' не найден.");
                return null;
            }

            // --- 4. Вставка фигур ---
            Visio.Shape? lastShape = null;
            const double MM_TO_INCH = 1.0 / 25.4;

            try
            {
                for (int i = 0; i < qty; i++)
                {
                    // Drop использует дюймы
                    lastShape = page.Drop(mst, xMM * MM_TO_INCH, yMM * MM_TO_INCH);
                }
            }
            catch (Exception ex)
            {
                // Логирование ошибки вставки, если нужно
                Console.WriteLine($"Ошибка при Drop: {ex.Message}");
            }

            return lastShape;
        }

        // ВСПОМОГАТЕЛЬНЫЙ МЕТОД ДЛЯ СБОРКИ ДАННЫХ
        // Обновленная сигнатура: принимает List<RawExcelHit> вместо Dictionary
        private void ProcessPageVba(Visio.Document doc, Visio.Page page, List<RawExcelHit> hits, VisioConfiguration config)
        {
            var itemsToDraw = new List<VisioItem>();

            // 1. ФИКСИРОВАННЫЕ ФИГУРЫ (Manual) - остаются без изменений
            if (config.PredefinedMasterConfigs != null)
            {
                foreach (var pm in config.PredefinedMasterConfigs)
                {
                    if (string.IsNullOrWhiteSpace(pm.MasterName)) continue;

                    // Парсинг координат с поддержкой точки и запятой
                    double xMM = 0, yMM = 0;
                    if (!string.IsNullOrWhiteSpace(pm.CoordinatesXY))
                    {
                        var parts = pm.CoordinatesXY.Split(',');
                        if (parts.Length >= 2)
                        {
                            double.TryParse(parts[0].Trim().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out xMM);
                            double.TryParse(parts[1].Trim().Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out yMM);
                        }
                    }

                    for (int i = 0; i < pm.Quantity; i++)
                    {
                        itemsToDraw.Add(new VisioItem
                        {
                            MasterName = pm.MasterName.Trim(),
                            X = xMM,
                            Y = yMM,
                            Anchor = pm.Anchor ?? "Center",
                            PlacementType = "Manual"
                        });
                    }
                }
            }

            // 2. НАЙДЕННЫЕ ФИГУРЫ (Sequential)
            if (hits != null && config.SequentialDrawing.Enabled)
            {
                // Группируем результаты поиска, чтобы не рисовать дубликаты, если это не нужно
                // Группируем по SearchTerm, так как именно он привязан к MasterName в конфиге
                var groupedHits = hits
                    .Where(h => h.ConditionMet) // Только те, где условие выполнено
                    .GroupBy(h => h.SearchTerm);

                foreach (var group in groupedHits)
                {
                    string searchTerm = group.Key;

                    // Находим правило в конфиге для этого поискового слова
                    var rule = config.SearchRules.FirstOrDefault(r =>
                        string.Equals(r.ExcelValue, searchTerm, StringComparison.OrdinalIgnoreCase));

                    // Если правило найдено и у него есть имя мастера Visio
                    if (rule != null && !string.IsNullOrWhiteSpace(rule.VisioMasterName))
                    {
                        // Считаем количество
                        int totalQty = 0;

                        // Проверяем каждое попадание в группе
                        foreach (var hit in group)
                        {
                            if (hit.IsLimited)
                                totalQty += 1; // Если ограничено - считаем как 1 (но тут нюанс: ограничение обычно на группу)
                            else
                                totalQty += hit.Quantity;
                        }

                        // КОРРЕКЦИЯ ЛОГИКИ LIMIT: 
                        // Если в правиле стоит LimitQuantity, то мы должны нарисовать фигуру ТОЛЬКО ОДИН РАЗ,
                        // независимо от того, сколько строк мы нашли.
                        if (rule.LimitQuantity)
                        {
                            totalQty = 1;
                        }

                        for (int k = 0; k < totalQty; k++)
                        {
                            itemsToDraw.Add(new VisioItem
                            {
                                MasterName = rule.VisioMasterName.Trim(),
                                X = 0, // Игнорируется для Sequential
                                Y = 0,
                                PlacementType = "Sequential",
                                Anchor = config.SequentialDrawing.Anchor
                            });
                        }
                    }
                }
            }

            // Запуск макроса
            VisioVbaRunner.RunDrawingMacro(doc, page, itemsToDraw, config);
        }

        /// <summary>
        /// Безопасно освобождает COM-объект, вызывая Marshal.ReleaseComObject в цикле, 
        /// пока счетчик ссылок не станет равным нулю.
        /// Это критически важно для корректного завершения работы с Visio Interop.
        /// </summary>
        private void ReleaseComObject(object? obj)
        {
            // Проверяем, что объект существует и является COM-объектом
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {
                    // Выполняем цикл для гарантированного освобождения.
                    // Если объект был передан нескольким переменным, Marshal.ReleaseComObject
                    // возвращает счетчик ссылок, который должен стать 0 для полного освобождения.
                    while (Marshal.ReleaseComObject(obj) > 0)
                    {
                        Marshal.ReleaseComObject(obj);
                    }
                }
                catch (Exception ex)
                {
                    // Здесь можно добавить логгирование, если не удается освободить объект
                    System.Diagnostics.Debug.WriteLine($"Ошибка при освобождении COM-объекта: {ex.Message}");
                }
                finally
                {
                    // Обнуляем ссылку в управляемом коде
                    obj = null;
                }
            }
        }

        /// <summary>
        /// Сопоставляет "Наименование" из Excel с ключом Visio Master из MasterMap.
        /// </summary>
        private List<Dictionary<string, string>> PrepareVisioData(
    List<Dictionary<string, string>> extractedData,
    List<SearchRule> SearchRules)
        {
            // ИСПРАВЛЕНО: используем аргумент SearchRules вместо несуществующего config.SearchRules
            var masterMap = RulesToMap(SearchRules);

            var visioData = new List<Dictionary<string, string>>();
            int totalItems = extractedData.Count;
            int mappedItems = 0;

            if (!masterMap.Any())
            {
                UpdateStatus("?? MasterMap пуст! Невозможно сопоставить данные.");
                return visioData;
            }

            UpdateStatus($"Начало сопоставления {totalItems} позиций Excel с MasterMap...");

            foreach (var item in extractedData)
            {
                if (!item.TryGetValue("Наименование", out string? content) || string.IsNullOrEmpty(content))
                    continue;

                string cleanedContent = content.Trim();
                bool matched = false;

                // Ищем самое длинное совпадение
                var bestMatch = masterMap.Keys
                    .Where(key => !string.IsNullOrEmpty(key))
                    .OrderByDescending(key => key.Length)
                    .FirstOrDefault(key =>
                        cleanedContent.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0);

                if (bestMatch != null)
                {
                    string visioMasterName = masterMap[bestMatch];
                    mappedItems++;
                    matched = true;

                    var newItem = new Dictionary<string, string>(item)
                    {
                        ["VisioMasterName"] = visioMasterName,
                        ["MatchedKey"] = bestMatch
                    };

                    visioData.Add(newItem);
                    UpdateStatus($"  ? Сопоставлено: '{bestMatch}' -> '{visioMasterName}'");
                }

                if (!matched)
                {
                    UpdateStatus($"  ? Не сопоставлено: '{cleanedContent}'");
                }
            }

            UpdateStatus($"? Сопоставление завершено. Найдено {mappedItems} из {totalItems} позиций.");
            return visioData;
        }

        // Этот метод заменяет старый GenerateVisioDocument и реализует логику
        // создания единого документа с двумя страницами и не закрывает Visio.
        // ИСПРАВЛЕНИЕ CS1501: Сигнатура метода теперь принимает 3 аргумента, как в OpenVisioClick
        // ИСПРАВЛЕНИЕ CS1501: Сигнатура метода теперь принимает 3 аргумента, как в OpenVisioClick

        private Dictionary<string, string> RulesToMap(List<SearchRule> rules)
        {
            return rules?
                .Where(r =>
                    !string.IsNullOrWhiteSpace(r.ExcelValue) &&
                    !string.IsNullOrWhiteSpace(r.VisioMasterName))
                .ToDictionary(
                    r => r.ExcelValue.Trim(),
                    r => r.VisioMasterName.Trim(),
                    StringComparer.OrdinalIgnoreCase
                )
                ?? new Dictionary<string, string>();
        }

        private void CreateUnifiedVisioFile(List<Dictionary<string, string>> extractedData,
                            VisioConfiguration configMarking,
                            VisioConfiguration configScheme)
        {
            Visio.Application? visioApp = null;
            Visio.Document? newDocument = null;

            try
            {
                UpdateStatus("?? Начало генерации объединенного Visio-файла...");

                // ИСПРАВЛЕНО: 
                // 1. Убраны лишние вызовы RulesToMap, так как PrepareVisioData делает это внутри.
                // 2. Передаем extractedData первым параметром.
                // 3. Передаем конкретные списки правил (configMarking.SearchRules) вторым параметром.

                var markingData = PrepareVisioData(extractedData, configMarking.SearchRules);
                var schemeData = PrepareVisioData(extractedData, configScheme.SearchRules);

                // 2. ЗАПУСК VISIO И СОЗДАНИЕ ДОКУМЕНТА
                visioApp = new Visio.Application();
                visioApp.Visible = true; // Оставляем Visio открытым

                newDocument = visioApp.Documents.Add(""); // Создаем новый документ

                // 3. СОЗДАНИЕ СТРАНИЦ
                Visio.Page pageMarking = newDocument.Pages[1];
                pageMarking.Name = "Маркировка";
                pageMarking.PageSheet.CellsU["PageUnits"].FormulaU = "8"; // METER(1)
                pageMarking.PageSheet.CellsU["DrawingUnits"].FormulaU = "8"; // METER(1)

                Visio.Page pageScheme = newDocument.Pages.Add();
                pageScheme.Name = "Схема";
                pageScheme.PageSheet.CellsU["PageUnits"].FormulaU = "8";
                pageScheme.PageSheet.CellsU["DrawingUnits"].FormulaU = "8";

                // 4. ЗАПОЛНЕНИЕ СТРАНИЦ
                PopulateVisioPage(pageMarking, markingData, configMarking, false);
                PopulateVisioPage(pageScheme, schemeData, configScheme, true);

                // 5. АКТИВИРУЕМ СХЕМУ ДЛЯ ПРОСМОТРА
                if (visioApp.ActiveWindow != null)
                {
                    visioApp.ActiveWindow.Page = pageScheme;
                }

                UpdateStatus($"? Файл Visio успешно сгенерирован и открыт");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                UpdateStatus($"? COM Ошибка Visio: {ex.Message}");
                MessageBox.Show($"COM Ошибка Visio: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                UpdateStatus($"? Общая ошибка при генерации Visio: {ex.Message}");
                MessageBox.Show($"Общая ошибка при генерации Visio: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void OpenAllStencils(Visio.Application visioApp, Visio.Document visioDoc, IEnumerable<string> stencilPaths, List<string> allFoundMasters)
        {
            var stencils = visioApp.Documents;

            foreach (string stencilPath in stencilPaths)
            {
                if (!File.Exists(stencilPath)) continue;

                Visio.Document? stencilDoc = null;
                try
                {
                    // Открываем трафарет скрытно
                    stencilDoc = stencils.OpenEx(stencilPath, (short)Visio.VisOpenSaveArgs.visOpenHidden);

                    // Собираем всех мастеров
                    foreach (Visio.Master master in stencilDoc.Masters)
                    {
                        // ИСПРАВЛЕНИЕ: Проверяем наличие с учетом регистра
                        if (!allFoundMasters.Contains(master.Name, StringComparer.OrdinalIgnoreCase))
                        {
                            allFoundMasters.Add(master.Name);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при открытии трафарета {stencilPath}: {ex.Message}");
                }
                finally
                {
                    // Освобождаем COM-объект трафарета
                    if (stencilDoc != null) Marshal.ReleaseComObject(stencilDoc);
                }
            }
            if (stencils != null) Marshal.ReleaseComObject(stencils);
        }

        // Новый вспомогательный метод, реализующий логику размещения фигур
        private void PopulateVisioPage(Visio.Page page, List<Dictionary<string, string>> extractedData,
                              VisioConfiguration config, bool isScheme)
        {
            // ИСПРАВЛЕНО: SPACING теперь используется в мм, а затем конвертируется
            const double SPACING_MM = 50.0; // 50 мм между фигурами (5 см)
            const double INITIAL_OFFSET_MM = 25.4;
            double SPACING_INCH = SPACING_MM * MM_TO_INCH; // Конвертируем в дюймы

            // Удалите const double PAGE_WIDTH = 0.297; - она не используется для настройки листа здесь

            // ИСПРАВЛЕНО: Начальные координаты в дюймах
            double initialOffsetInches = INITIAL_OFFSET_MM * MM_TO_INCH; // 1 дюйм от края
            double currentX = initialOffsetInches;
            double currentY = initialOffsetInches;

            var openStencils = new List<Visio.Document>();
            var mastersNotFound = new List<string>(); // Для сбора ненайденных мастеров
            var allAvailableMasterNames = new HashSet<string>(StringComparer.Ordinal); // Для сбора всех доступных имен

            Visio.Master? master = null;
            Visio.Shape? shape = null;

            try
            {
                UpdateStatus($"Начало размещения фигур на странице '{page.Name}'. Попытка открыть трафареты...");

                // 1. ОТКРЫТИЕ ВСЕХ ТРАФАРЕТОВ (БЕЗ ИЗМЕНЕНИЙ)
                foreach (string path in config.StencilFilePaths)
                {
                    if (string.IsNullOrEmpty(path)) continue;
                    try
                    {
                        if (!System.IO.File.Exists(path))
                        {
                            UpdateStatus($"? ОШИБКА: Файл трафарета не существует: {path}");
                            continue;
                        }

                        Visio.Document stencilDoc = page.Application.Documents.Open(path);
                        openStencils.Add(stencilDoc);
                        UpdateStatus($"? Трафарет открыт: {Path.GetFileName(path)}");
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"? КРИТИЧЕСКАЯ ОШИБКА открытия трафарета '{Path.GetFileName(path)}': {ex.Message}");
                    }
                }

                if (!openStencils.Any())
                {
                    UpdateStatus("? КРИТИЧЕСКАЯ ОШИБКА: Не удалось открыть ни один трафарет Visio. Прерывание.");
                    return;
                }

                // 2. СБОР ВСЕХ ДОСТУПНЫХ МАСТЕРОВ (для диагностики)
                foreach (var stencilDoc in openStencils)
                {
                    foreach (Visio.Master m in stencilDoc.Masters)
                    {
                        // Собираем точные имена (NameU)
                        allAvailableMasterNames.Add(m.NameU);
                        ReleaseComObject(m); // Освобождаем мастер немедленно
                    }


                }

                UpdateStatus($"Найдено {allAvailableMasterNames.Count} уникальных фигур в открытых трафаретах.");

                // =========================================================================
                // 3. ДОБАВЛЕНИЕ ПРЕДОПРЕДЕЛЕННЫХ ФИГУР (НОВАЯ ЛОГИКА)
                // =========================================================================
                // ... внутри PopulateVisioPage ...

                // =========================================================================
                // 3. ДОБАВЛЕНИЕ ПРЕДОПРЕДЕЛЕННЫХ ФИГУР (НОВАЯ ЛОГИКА)
                // =========================================================================
                UpdateStatus($"Добавление {config.PredefinedMasterConfigs.Count} предопределенных фигур...");

                // ИСПРАВЛЕНО: тип переменной цикла изменен на var (PredefinedMasterConfig)
                foreach (var pmConfig in config.PredefinedMasterConfigs.Where(n => !string.IsNullOrWhiteSpace(n.MasterName)))
                {
                    master = null;
                    shape = null;

                    // ИСПРАВЛЕНО: берем имя из объекта конфига
                    string predefinedMasterName = pmConfig.MasterName;

                    // 3.1. Поиск мастера
                    foreach (var stencilDoc in openStencils)
                    {
                        try
                        {
                            master = stencilDoc.Masters.Cast<Visio.Master>().FirstOrDefault(m =>
                                m.NameU.Equals(predefinedMasterName, StringComparison.Ordinal));

                            if (master != null) break;
                        }
                        catch (Exception)
                        {
                            // Просто игнорируем, если мастер не найден в этом трафарете
                        }
                    }

                    if (master != null)
                    {
                        try
                        {
                            // 3.2. Добавление фигуры: КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: Используем координаты из конфига
                            double xDrop = currentX, yDrop = currentY; // Fallback к автоматическому размещению
                            bool usedAutoPlacement = true;

                            // Проверяем, заданы ли координаты в конфиге
                            if (!string.IsNullOrWhiteSpace(pmConfig.CoordinatesXY))
                            {
                                var parts = pmConfig.CoordinatesXY.Split(',');
                                if (parts.Length >= 2 &&
                                    float.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float xMM) &&
                                    float.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float yMM))
                                {
                                    // Конвертация из MM в INCH
                                    xDrop = xMM * MM_TO_INCH;
                                    yDrop = yMM * MM_TO_INCH;
                                    usedAutoPlacement = false;
                                }
                            }

                            shape = page.Drop(master, xDrop, yDrop); // Используем конвертированные или автоматические координаты
                            shape.Text = $"Предопределенная: {predefinedMasterName}";

                            // Логика автоматического смещения (только если не использовались координаты из конфига)
                            if (usedAutoPlacement)
                            {
                                currentX += SPACING_INCH;
                                if (currentX > 10.0)
                                {
                                    currentX = initialOffsetInches;
                                    currentY += SPACING_INCH;
                                }
                            }

                            UpdateStatus($"  ? Добавлена предопределенная фигура: '{predefinedMasterName}'");
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"? Ошибка размещения предопределенного мастера '{predefinedMasterName}': {ex.Message}");
                        }
                    }
                    else
                    {
                        // Если предопределенный мастер не найден
                        if (!mastersNotFound.Contains(predefinedMasterName))
                        {
                            mastersNotFound.Add(predefinedMasterName);
                        }
                    }

                    // Освобождаем объекты, созданные в этом цикле
                    ReleaseComObject(shape);
                    ReleaseComObject(master);
                }

                // 3. РАЗМЕЩЕНИЕ ФИГУР НА СТРАНИЦЕ
                foreach (var dataItem in extractedData)
                {
                    if (!dataItem.ContainsKey("VisioMasterName") ||
                        !int.TryParse(dataItem.GetValueOrDefault("Количество", "0"), out int count)) continue;

                    string masterName = dataItem["VisioMasterName"];
                    master = null;

                    // 3.1. БЫСТРАЯ ПРОВЕРКА (Она не гарантирует, что Visio найдет объект, но дает нам точную диагностику)
                    if (!allAvailableMasterNames.Contains(masterName))
                    {
                        if (!mastersNotFound.Contains(masterName)) mastersNotFound.Add(masterName);
                        UpdateStatus($"? Мастер '{masterName}' отсутствует в списке доступных фигур.");
                        continue;
                    }

                    // 3.2. ПОИСК ИЗ КОНКРЕТНОГО ТРАФАРЕТА (если имя мастера существует)
                    foreach (var stencilDoc in openStencils)
                    {
                        // Имя мастера должно ТОЧНО совпадать с NameU
                        master = stencilDoc.Masters.Cast<Visio.Master>().FirstOrDefault(m =>
                            m.NameU.Equals(masterName, StringComparison.Ordinal));

                        if (master != null)
                        {
                            break;
                        }
                    }

                    // 3.3. ОБРАБОТКА СЛУЧАЯ, КОГДА МАСТЕР НЕ НАЙДЕН (ХОТЯ ИМЯ БЫЛО В СПИСКЕ)
                    if (master == null)
                    {
                        // Это маловероятная ошибка (значит COM-объект не может быть получен)
                        if (!mastersNotFound.Contains(masterName)) mastersNotFound.Add(masterName);
                        UpdateStatus($"? Не удалось получить COM-объект мастера '{masterName}', хотя он был найден в списке.");
                        continue;
                    }

                    // 3.4. РАЗМЕЩЕНИЕ ФИГУР
                    try
                    {
                        for (int i = 0; i < count; i++)
                        {
                            shape = page.Drop(master, currentX, currentY); // currentX/Y теперь в дюймах
                            shape.Text = masterName;

                            // Логика смещения
                            currentX += SPACING_INCH; // ИСПРАВЛЕНО: используем SPACING_INCH
                            if (currentX > 10.0)
                            {
                                currentX = initialOffsetInches;
                                currentY += SPACING_INCH; // ИСПРАВЛЕНО: используем SPACING_INCH
                            }

                            ReleaseComObject(shape);
                            shape = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"? Ошибка размещения мастера '{masterName}': {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObject(master);
                    }
                }

                UpdateStatus($"? Размещение фигур на странице '{page.Name}' завершено.");

                // 4. ВЫВОД СООБЩЕНИЯ О НЕЙДЕННЫХ МАСТЕРАХ
                if (mastersNotFound.Any())
                {
                    string missingMastersList = string.Join(Environment.NewLine, mastersNotFound.Distinct());

                    // Ограничиваем список доступных имен для удобства чтения
                    string availableList = allAvailableMasterNames.Any()
                        ? string.Join(", ", allAvailableMasterNames.OrderBy(n => n).Take(50))
                        : "Нет доступных фигур.";

                    MessageBox.Show(
                        this,
                        "? КРИТИЧЕСКАЯ ОШИБКА: Не удалось разместить следующие фигуры (Мастера):\n\n" +
                        $"ИСКОМЫЕ ФИГУРЫ:\n{missingMastersList}\n\n" +
                        "==========================================================\n" +
                        "ДОСТУПНЫЕ ФИГУРЫ В ТРАФАРЕТАХ (Master.NameU):\n" +
                        $"{availableList}" +
                        (allAvailableMasterNames.Count > 50 ? $"\n... всего {allAvailableMasterNames.Count} фигур." : "") +
                        "\n\nПРИМЕЧАНИЕ: Искомые имена должны ТОЧНО совпадать с доступными именами (включая регистр!).",
                        $"Проблема с фигурами на странице '{page.Name}'",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            finally
            {
                // 5. ОЧИСТКА COM-ОБЪЕКТОВ
                foreach (var stencilDoc in openStencils)
                {
                    try
                    {
                        stencilDoc.Close();
                    }
                    catch { /* Игнорируем ошибки при закрытии */ }
                    finally
                    {
                        ReleaseComObject(stencilDoc);
                    }
                }
            }
        }

        private Visio.Master? GetMasterFromStencilOrDocument(Visio.Application visioApp, Visio.Document visioDoc, List<string> stencilFilePaths, string masterName)
        {
            Visio.Master? masterToUse = null;

            foreach (var doc in visioApp.Documents)
            {
                if (doc is Visio.Document stencilDoc)
                {
                    try
                    {
                        masterToUse = stencilDoc.Masters.get_ItemU(masterName);
                        if (masterToUse != null) return masterToUse;
                    }
                    catch { }
                }
            }

            foreach (string stencilPath in stencilFilePaths)
            {
                if (!System.IO.File.Exists(stencilPath)) continue;

                Visio.Document? stencilDoc = null;
                try
                {
                    stencilDoc = visioApp.Documents.OpenEx(stencilPath, (short)Visio.VisOpenSaveArgs.visOpenDocked | (short)Visio.VisOpenSaveArgs.visOpenHidden);
                    masterToUse = stencilDoc.Masters.get_ItemU(masterName);

                    if (masterToUse != null)
                    {
                        return masterToUse;
                    }
                }
                catch { }
                finally
                {
                    if (masterToUse == null && stencilDoc != null)
                    {
                        try { stencilDoc.Close(); } catch { }
                        Marshal.ReleaseComObject(stencilDoc);
                    }
                }
            }


            if (masterToUse == null && masterName == "Rectangle")
            {
                try
                {
                    masterToUse = visioDoc.Masters.get_ItemU("Rectangle");
                }
                catch { }
            }

            return masterToUse;
        }

        // --- ВСТАВИТЬ ПЕРЕД КОНЦОМ КЛАССА FORM1 ---

        // 1. Функция чтения Excel и превращения его в список объектов
        private List<EquipmentItem> ParseExcelData(ExcelWorksheet sheet)
        {
            var foundEquipment = new List<EquipmentItem>();

            // ВАЖНО: Проверьте номера строк и колонок под ваш файл!
            int startRow = 10; // С какой строки начинаются данные
            int endRow = sheet.Dimension?.End.Row ?? 100;

            for (int row = startRow; row <= endRow; row++)
            {
                // Предположим, наименование в 1-й колонке, кол-во в 5-й (поправьте индексы!)
                string excelName = sheet.Cells[row, 1].Text.Trim();
                string qtyText = sheet.Cells[row, 5].Text.Trim();

                if (string.IsNullOrEmpty(excelName)) continue;

                // Ищем совпадение в нашем словаре настроек
                if (_equipmentConfig.TryGetValue(excelName, out EquipmentItem configItem))
                {
                    int.TryParse(qtyText, out int qty);
                    if (qty > 0)
                    {
                        // Создаем копию объекта с реальным количеством
                        foundEquipment.Add(new EquipmentItem
                        {
                            OriginalName = excelName,
                            ShortName = configItem.ShortName,
                            // ИСПРАВЛЕНО: Сначала закрываем скобку Count(...), а потом прибавляем + 1
                            PositionCode = configItem.PositionCode + (foundEquipment.Count(x => x.PositionCode != null && x.PositionCode.StartsWith(configItem.PositionCode)) + 1),
                            ShapeMasterName = configItem.ShapeMasterName,
                            Quantity = qty
                        });
                    }
                }
            }
            return foundEquipment;
        }

        // 2. Функция записи данных внутрь фигуры Visio (Shape Data)
        private void UpdateVisioShapeData(Visio.Shape shape, EquipmentItem item)
        {
            try
            {
                // Пишем текст на фигуре
                shape.Text = $"{item.PositionCode}\n{item.ShortName}";

                // Проверяем наличие секции свойств
                if (shape.get_SectionExists((short)Visio.VisSectionIndices.visSectionProp, 0) == 0)
                    shape.AddSection((short)Visio.VisSectionIndices.visSectionProp);

                // Записываем скрытые данные (для отчетов Visio)
                SetShapeProperty(shape, "ShortName", "Наименование", item.ShortName);
                SetShapeProperty(shape, "Position", "Позиция", item.PositionCode);
                SetShapeProperty(shape, "Quantity", "Кол-во", item.Quantity.ToString());
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("Ошибка обновления фигуры: " + ex.Message); }
        }

        // 3. Вспомогательная функция для свойств
        private void SetShapeProperty(Visio.Shape shape, string propName, string label, string value)
        {
            string cellName = "Prop." + propName;
            if (shape.get_CellExists(cellName, (short)Visio.VisExistsFlags.visExistsAnywhere) == 0)
                shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionProp, propName, (short)Visio.VisRowTags.visTagDefault);

            shape.CellsU[cellName].FormulaU = "\"" + value + "\"";
            shape.CellsU[cellName + ".Label"].FormulaU = "\"" + label + "\"";
        }

        // 4. Функция рисования ТАБЛИЦЫ
        private void DrawSpecificationTable(Visio.Page page, List<EquipmentItem> items)
        {
            double x = 1.0; // Отступ слева (дюймы)
            double y = 10.0; // Отступ сверху (дюймы, Visio считает от нижнего края, поэтому 10 - это высоко)
            double rowHeight = 0.25;

            // Заголовки
            page.DrawRectangle(x, y, x + 0.5, y + rowHeight).Text = "Поз.";
            page.DrawRectangle(x + 0.5, y, x + 2.5, y + rowHeight).Text = "Наименование";
            page.DrawRectangle(x + 2.5, y, x + 3.0, y + rowHeight).Text = "Кол.";

            y -= rowHeight;

            foreach (var item in items)
            {
                page.DrawRectangle(x, y, x + 0.5, y + rowHeight).Text = item.PositionCode;
                page.DrawRectangle(x + 0.5, y, x + 2.5, y + rowHeight).Text = item.ShortName;
                page.DrawRectangle(x + 2.5, y, x + 3.0, y + rowHeight).Text = item.Quantity.ToString();
                y -= rowHeight;
            }
        }
    }



    // =========================================================================
    // 5. НОВАЯ ФОРМА ВЫБОРА ЛИСТОВ (SheetSelectionForm)
    // =========================================================================

    public class SheetSelectionForm : Form
    {
        private readonly CheckedListBox _clbSheets;
        public List<string> SelectedSheets { get; private set; } = new List<string> { "1.ТЗ на объект ZONT для VISIO" };

        public SheetSelectionForm(List<string> allSheetNames, List<string> initialSelectedSheets)
        {
            this.Text = "Выбор листов для анализа";
            this.Size = new Size(400, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 2,
                RowStyles =
                {
                    new RowStyle(SizeType.Percent, 100),
                    new RowStyle(SizeType.Absolute, 50)
                }
            };
            this.Controls.Add(mainLayout);

            _clbSheets = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                SelectionMode = SelectionMode.One // <-- ИСПРАВЛЕНО: Установка SelectionMode в One для корректной работы кликов по элементам.
            };
            mainLayout.Controls.Add(_clbSheets, 0, 0);

            // Заполняем список и устанавливаем начальные галочки
            foreach (var sheetName in allSheetNames)
            {
                // Проверяем, был ли этот лист выбран ранее (без учета регистра)
                bool isChecked = initialSelectedSheets.Contains(sheetName, StringComparer.OrdinalIgnoreCase);
                _clbSheets.Items.Add(sheetName, isChecked);
            }

            // Кнопки
            var footerFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(5)
            };
            mainLayout.Controls.Add(footerFlow, 0, 1);

            var btnOk = new Button { Text = "ОК", Width = 100, Height = 30, DialogResult = DialogResult.OK };
            btnOk.Click += (s, e) =>
            {
                // Собираем выбранные листы
                SelectedSheets = _clbSheets.CheckedItems.Cast<string>().ToList();
            };

            var btnCancel = new Button { Text = "Отмена", Width = 100, Height = 30, DialogResult = DialogResult.Cancel };

            footerFlow.Controls.Add(btnCancel);
            footerFlow.Controls.Add(btnOk);
        }
    }

}



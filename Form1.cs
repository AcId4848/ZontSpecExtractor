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


namespace ZontSpecExtractor
{
    // =========================================================================
    // 1. КОНФИГУРАЦИЯ И НАСТРОЙКИ (Сохранение состояния)
    // =========================================================================

    // Вне класса VisioConfiguration (но в том же namespace)
    public record PredefinedMasterConfig
    {
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

        // Свойство для совместимости, если где-то используется old-style "Term"
        public string Term
        {
            get => ExcelValue;
            set => ExcelValue = value;
        }
    }



    public class SearchConfiguration
    {
        public List<string> TargetSheetNames { get; set; } = new List<string> { "1.ТЗ на объект ZONT" };

        // Заменяем простой список строк на список правил
        public List<SearchRule> Rules { get; set; } = new List<SearchRule>();
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
            btnSave.Click += BtnSave_Click;
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
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 600 }; // Расширили левую часть

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

            // 2. Группа: Найденные фигуры (Sequential) - НОВОЕ ОКНО
            var grpSeq = new GroupBox { Text = "Настройки для НАЙДЕННЫХ фигур (поток)", Dock = DockStyle.Top, Height = 140 };
            var flowSeq = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };

            // Поля ввода
            var txtStart = new TextBox { Text = cfg.SequentialDrawing.StartCoordinatesXY, Width = 80 };

            // ИСПРАВЛЕНИЕ: Устанавливаем Minimum и Maximum перед Value
            var numMaxW = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 1000,
                Value = (decimal)cfg.SequentialDrawing.MaxLineWidthMM,
                Width = 60
            };
            var numVGap = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Value = (decimal)cfg.SequentialDrawing.VerticalStepMM,
                Width = 60
            };
            var numHGap = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Value = (decimal)cfg.SequentialDrawing.HorizontalStepMM,
                Width = 60
            };

            var chkEn = new CheckBox { Text = "Включить", Checked = cfg.SequentialDrawing.Enabled, Width = 80 };

            // Добавляем контролы с подписями
            AddControl(flowSeq, "Старт X,Y (мм):", txtStart);
            AddControl(flowSeq, "Макс. ширина строки (мм):", numMaxW);
            AddControl(flowSeq, "Отступ снизу (мм):", numVGap);
            AddControl(flowSeq, "Отступ сбоку (мм):", numHGap);
            flowSeq.Controls.Add(chkEn);

            grpSeq.Controls.Add(flowSeq);
            leftPanel.Controls.Add(grpSeq);

            // 3. Таблица фиксированных фигур
            var gridLabel = new Label { Text = "Фиксированные фигуры (Шапки, Рамки):", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9, FontStyle.Bold), Padding = new Padding(0, 5, 0, 0) };
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                ColumnHeadersHeight = 30
            };
            dgv.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Имя мастера", DataPropertyName = "MasterName", Width = 150 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "X (мм)", DataPropertyName = "X", Width = 50 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Y (мм)", DataPropertyName = "Y", Width = 50 });

            var colAnchor = new DataGridViewComboBoxColumn { HeaderText = "Anchor", DataPropertyName = "Anchor", Width = 80 };
            colAnchor.Items.AddRange("Center", "TopLeft", "BottomLeft");
            dgv.Columns.Add(colAnchor);

            var bindingList = new System.ComponentModel.BindingList<PredefinedViewModel>(
                cfg.PredefinedMasterConfigs.Select(p => new PredefinedViewModel(p)).ToList()
            );
            dgv.DataSource = bindingList;

            // Сборка левой панели (Порядок добавления важен для Dock)
            leftPanel.Controls.Add(dgv);     // Fill
            leftPanel.Controls.Add(gridLabel); // Top
            leftPanel.Controls.Add(grpSeq);  // Top
            leftPanel.Controls.Add(topPanel); // Top
            split.Panel1.Controls.Add(leftPanel);

            // --- ПРАВАЯ ЧАСТЬ: ПРЕВЬЮ ---
            var previewPanel = new PagePreviewPanel
            {
                Dock = DockStyle.Fill,
                PageSize = cfg.PageSize,
                Orientation = cfg.PageOrientation,
                FixedShapes = cfg.PredefinedMasterConfigs, // <--- ИСПРАВЛЕНИЕ: Используем FixedShapes
                SeqConfig = cfg.SequentialDrawing
            };
            split.Panel2.Controls.Add(previewPanel);

            // --- СОБЫТИЯ ---
            Action updatePreview = () =>
            {
                // Обновляем конфиг из UI
                cfg.PageSize = cbSize.Text;
                cfg.SequentialDrawing.StartCoordinatesXY = txtStart.Text;
                cfg.SequentialDrawing.MaxLineWidthMM = (double)numMaxW.Value;
                cfg.SequentialDrawing.VerticalStepMM = (double)numVGap.Value;
                cfg.SequentialDrawing.HorizontalStepMM = (double)numHGap.Value;
                cfg.SequentialDrawing.Enabled = chkEn.Checked;

                // Обновляем список предопределенных
                cfg.PredefinedMasterConfigs = bindingList.Select(vm => vm.ToConfig()).ToList();

                // Обновляем свойства панели
                previewPanel.PageSize = cfg.PageSize;
                previewPanel.Orientation = cfg.PageOrientation;
                previewPanel.FixedShapes = cfg.PredefinedMasterConfigs; // Обновляем ссылку
                previewPanel.SeqConfig = cfg.SequentialDrawing;

                previewPanel.Invalidate(); // Перерисовка
            };

            // Привязываем события
            cbSize.SelectedIndexChanged += (s, e) => updatePreview();
            btnOrient.Click += (s, e) => {
                cfg.PageOrientation = (cfg.PageOrientation == "Portrait") ? "Landscape" : "Portrait";
                btnOrient.Text = cfg.PageOrientation == "Landscape" ? "Альбомная" : "Книжная";
                previewPanel.Orientation = cfg.PageOrientation; // Сразу меняем в панели
                updatePreview();
            };

            txtStart.TextChanged += (s, e) => updatePreview();
            numMaxW.ValueChanged += (s, e) => updatePreview();
            numVGap.ValueChanged += (s, e) => updatePreview();
            numHGap.ValueChanged += (s, e) => updatePreview();
            chkEn.CheckedChanged += (s, e) => updatePreview();

            dgv.CellEndEdit += (s, e) => updatePreview();
            dgv.RowsRemoved += (s, e) => updatePreview();
            dgv.UserAddedRow += (s, e) => updatePreview();

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
                RowCount = 4,
                RowStyles = {
            new RowStyle(SizeType.Absolute, 30),
            new RowStyle(SizeType.Absolute, 60),
            new RowStyle(SizeType.Absolute, 30),
            new RowStyle(SizeType.Percent, 100)
        }
            };

            // 1. Листы
            layout.Controls.Add(new Label { Text = "Целевые листы (откуда брать данные):", Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);
            _rtxtSheetNames = new RichTextBox { Dock = DockStyle.Fill, Text = string.Join(Environment.NewLine, AppSettings.SearchConfig.TargetSheetNames), ReadOnly = true, BackColor = SystemColors.ControlLight };
            layout.Controls.Add(_rtxtSheetNames, 0, 1);

            // 2. Правила поиска
            layout.Controls.Add(new Label { Text = "Правила поиска и условий:", Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 2);

            _dgvSearchRules = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                ColumnHeadersHeight = 40, // Чуть выше для читаемости
                AllowUserToResizeColumns = true
            };

            // --- КОЛОНКИ ТАБЛИЦЫ ---

            // 1. Искомое слово
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Искомое слово\n(Visio Key)",
                DataPropertyName = "ExcelValue",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 150
            });

            // 2. Где искать само слово (Например "C")
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

            // 4. Где искать условие (Например "L")
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Где условие\n(Col: L)",
                DataPropertyName = "ConditionColumn",
                Width = 80,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            // 5. Значение условия (Например "1")
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

            // Привязка данных
            _dgvSearchRules.DataSource = new System.ComponentModel.BindingList<SearchRule>(AppSettings.SearchConfig.Rules ?? new List<SearchRule>());

            layout.Controls.Add(_dgvSearchRules, 0, 3);

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
            new RowStyle(SizeType.Absolute, 20), // Заголовок
            new RowStyle(SizeType.Absolute, 20), // Label Трафареты
            new RowStyle(SizeType.Absolute, 60), // Путь к трафаретам
            new RowStyle(SizeType.Absolute, 35), // Кнопки
            new RowStyle(SizeType.Percent, 40),  // НОВОЕ: Окно найденных фигур (AvailableMasters)
            new RowStyle(SizeType.Absolute, 20), // Label Карта
            new RowStyle(SizeType.Percent, 60)   // Карта
        }
            };

            l.Controls.Add(new Label { Text = title, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);

            // 1. Трафареты
            l.Controls.Add(new Label { Text = "Трафареты:", Dock = DockStyle.Bottom }, 0, 1);
            rPath.Text = string.Join(Environment.NewLine, cfg.StencilFilePaths);
            l.Controls.Add(rPath, 0, 2);

            // 2. Кнопки
            var btnP = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            var btnAdd = new Button { Text = "+", Width = 30 };
            btnAdd.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Multiselect = true, Filter = "Visio|*.vssx;*.vsdx" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        cfg.StencilFilePaths.AddRange(ofd.FileNames.Where(f => !cfg.StencilFilePaths.Contains(f)));
                        rPath.Text = string.Join(Environment.NewLine, cfg.StencilFilePaths);
                    }
                }
            };
            btnP.Controls.Add(btnAdd);
            var btnClear = new Button { Text = "X", Width = 30, ForeColor = Color.Red };
            btnClear.Click += (s, e) => { cfg.StencilFilePaths.Clear(); rPath.Text = ""; };
            btnP.Controls.Add(btnClear);

            // Кнопка для сканирования, чтобы узнать имена
            var btnScan = new Button { Text = "Сканировать", AutoSize = true, BackColor = Color.LightGreen };
            // Мы передаем rFoundMasters, чтобы сразу обновить текст
            var rFoundMasters = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8), BackColor = Color.Beige };
            btnScan.Click += (s, e) => {
                ScanMasters(cfg); // Обновляет cfg.AvailableMasters
                rFoundMasters.Text = string.Join(Environment.NewLine, cfg.AvailableMasters.OrderBy(m => m)); // Обновляем окно
            };
            btnP.Controls.Add(btnAdd);
            btnP.Controls.Add(btnClear); // определите кнопку удаления как раньше
            btnP.Controls.Add(btnScan);
            l.Controls.Add(btnP, 0, 3);

            // 3. НОВОЕ ОКНО: Найденные фигуры (Сохраняется с настройками)
            // Загружаем сохраненные
            rFoundMasters.Text = string.Join(Environment.NewLine, cfg.AvailableMasters.OrderBy(m => m));
            // При изменении текста пользователем вручную - сохраняем обратно в конфиг
            rFoundMasters.TextChanged += (s, e) => {
                cfg.AvailableMasters = rFoundMasters.Text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
            };

            var grpFound = new GroupBox { Text = "Найденные фигуры (можно копировать):", Dock = DockStyle.Fill };
            grpFound.Controls.Add(rFoundMasters);
            l.Controls.Add(grpFound, 0, 4);

            // 4. Карта
            l.Controls.Add(new Label { Text = "Карта (Excel=Col=Master):", Dock = DockStyle.Bottom }, 0, 5);
            rMap.Text = string.Join(Environment.NewLine, cfg.SearchRules.Select(r => $"{r.ExcelValue}={r.SearchColumn}={r.VisioMasterName}"));
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
            if (rtb == null || string.IsNullOrWhiteSpace(rtb.Text)) return rules;

            // Разбиваем текст на строки
            var lines = rtb.Text.Split(new[] { Environment.NewLine, "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var trimLine = line.Trim();
                if (string.IsNullOrWhiteSpace(trimLine) || trimLine.StartsWith("//")) continue;

                // Разбиваем по знаку равенства
                var parts = trimLine.Split('=');

                // ВАРИАНТ 1: "Слово = Мастер" (2 части)
                if (parts.Length == 2)
                {
                    rules.Add(new SearchRule
                    {
                        ExcelValue = parts[0].Trim(),      // "Слово"
                        SearchColumn = "",                 // Колонка не указана
                        VisioMasterName = parts[1].Trim()  // "Мастер"
                    });
                }
                // ВАРИАНТ 2: "Слово = Колонка = Мастер" (3 части)
                else if (parts.Length == 3)
                {
                    rules.Add(new SearchRule
                    {
                        ExcelValue = parts[0].Trim(),      // "Слово"
                        SearchColumn = parts[1].Trim(),    // "Колонка" (например "V")
                        VisioMasterName = parts[2].Trim()  // "Мастер"
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
        public string FullItemName { get; set; } // Сюда теперь попадет "Насос ГВС", а не просто "Насос"
        public string SearchTerm { get; set; }
        public bool ConditionMet { get; set; }
        public int Quantity { get; set; }
        public bool IsLimited { get; set; }

        public SearchRule FoundRule { get; set; }
        public string TargetMasterName { get; set; } // <-- Для вывода в таблицу (Visio Name)
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
        private static readonly string[] COLS_OUT = {
            "Лист", "Наименование", "Количество"
        };

        private List<Dictionary<string, string>> data;
        private DataGridView dataGridView = null!;
        private Label lblFileInfo = null!;
        private Label lblStatus = null!;
        private Panel statusPanel = null!;
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

            UpdateStatus("⏳ Обновление размеров фигур из Visio...");

            // 3. ПАКЕТНОЕ ПОЛУЧЕНИЕ РАЗМЕРОВ (Один вызов Visio)
            var sizes = VisioScanner.GetMastersDimensionsBatch(allMasters, allStencils);

            // 4. Применяем размеры ко всем конфигам
            ApplySizesToConfig(AppSettings.SchemeConfig, sizes);
            ApplySizesToConfig(AppSettings.LabelingConfig, sizes);
            ApplySizesToConfig(AppSettings.CabinetConfig, sizes);

            UpdateStatus("✅ Размеры фигур обновлены.");
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
                UpdateStatus("⚠️ Не удалось найти ни одной фигуры Visio в указанных трафаретах.");
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

                    // Шаг 1: Получаем все имена листов из выбранного файла
                    var allSheetNames = GetSheetNames(AppSettings.LastLoadedFilePath);

                    if (!allSheetNames.Any())
                    {
                        MessageBox.Show("Не удалось получить имена листов. Проверьте файл.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Шаг 2: Открываем форму выбора листов, передавая текущие сохраненные листы как начальный выбор
                    using (var sheetForm = new SheetSelectionForm(allSheetNames, AppSettings.SearchConfig.TargetSheetNames))
                    {
                        if (sheetForm.ShowDialog() == DialogResult.OK)
                        {
                            // Шаг 3: Сохраняем выбранные листы в настройки
                            AppSettings.SearchConfig.TargetSheetNames = sheetForm.SelectedSheets;
                            AppSettings.Save(); // Сохраняем настройки в файл

                            // Шаг 4: Запускаем загрузку файлов с новыми настройками
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
                MessageBox.Show("Сначала необходимо загрузить и проанализировать Excel-файл (кнопка 'Загрузить Excel')!");
                return;
            }

            // 1. Группируем сырые данные по ДОСЛОВНОМУ названию позиции
            var detailedTableData = _rawHits
                .Where(h => h.ConditionMet) // Только те, которые прошли проверку условий (число в ячейке)
                .GroupBy(h => h.FullItemName) // <--- ГРУППИРУЕМ ПО ДОСЛОВНОМУ СОДЕРЖАНИЮ ЯЧЕЙКИ
                .Select((g, index) => new
                {
                    Number = index + 1,        // Номер по порядку
                    Position = g.Key,          // Позиция (дословно)
                    Count = g.Sum(x => x.Quantity) // Сумма количества
                })
                .OrderBy(x => x.Number)
                .ToList();

            if (!detailedTableData.Any())
            {
                MessageBox.Show("Не найдено позиций, удовлетворяющих правилам поиска и условиям.");
                return;
            }

            // 2. Создаем и показываем новую форму с таблицей
            var tableForm = new Form
            {
                Text = "Таблица найденных позиций (Спецификация)",
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent
            };

            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoGenerateColumns = true,
                DataSource = detailedTableData
            };

            // Переименование заголовков согласно требованиям:
            dgv.DataBindingComplete += (s, ev) =>
            {
                if (dgv.Columns["Number"] != null) dgv.Columns["Number"].HeaderText = "Номер по порядку";
                if (dgv.Columns["Position"] != null) dgv.Columns["Position"].HeaderText = "Позиция (дословно)";
                if (dgv.Columns["Count"] != null) dgv.Columns["Count"].HeaderText = "Количество";

                // Настройка ширины
                if (dgv.Columns["Number"] != null) dgv.Columns["Number"].Width = 100;
                if (dgv.Columns["Position"] != null) dgv.Columns["Position"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                if (dgv.Columns["Count"] != null) dgv.Columns["Count"].Width = 100;
            };

            tableForm.Controls.Add(dgv);
            tableForm.ShowDialog();
        }

        private async void LoadFiles(string[] filePaths)
        {
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

            
            UpdateStatus("⚙️ Идет сканирование Excel-файла...");
            await Task.Run(() =>
            {
                foreach (var path in filePaths)
                {
                    // ТЕПЕРЬ ТИПЫ СОВПАДАЮТ
                    var hits = ScanSpecificSheet(path);
                    _rawHits.AddRange(hits);
                }
            });

            // 1. Агрегируем сырые данные (_rawHits) в формат, нужный для Visio (data)
            data = GroupRawHitsForVisio(_rawHits); // <--- ВЫЗОВ НОВОЙ ФУНКЦИИ

            // 2. Обновляем основной DataGridView (который показывает агрегированные данные для Visio)
            UpdateDataGridView();
            ShowResultMessage(data.Count);
        }

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
        private List<RawExcelHit> ScanSpecificSheet(string filePath)
        {
            var rawHits = new List<RawExcelHit>();

            // Лицензия EPPlus
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Определяем листы для поиска
                    var targetSheets = AppSettings.SearchConfig.TargetSheetNames;
                    if (targetSheets == null || !targetSheets.Any())
                    {
                        targetSheets = package.Workbook.Worksheets.Select(w => w.Name).ToList();
                    }

                    foreach (var sheetName in targetSheets)
                    {
                        var ws = package.Workbook.Worksheets[sheetName];
                        if (ws == null || ws.Dimension == null) continue;

                        int startRow = ws.Dimension.Start.Row;
                        int endRow = ws.Dimension.End.Row;

                        // --- ЦИКЛ ПО СТРОКАМ EXCEL ---
                        for (int row = startRow; row <= endRow; row++)
                        {
                            // --- ЦИКЛ ПО ПРАВИЛАМ ПОИСКА ---
                            foreach (var rule in AppSettings.SearchConfig.Rules)
                            {
                                // =========================================================
                                // ЭТАП 0: ПРОВЕРКА ДОПОЛНИТЕЛЬНОГО УСЛОВИЯ (Condition)
                                // =========================================================
                                // Если галочка UseCondition стоит, мы проверяем соседнюю ячейку.
                                // Например: Колонка "L" должна быть равна "1".

                                if (rule.UseCondition)
                                {
                                    // Превращаем букву колонки (например "L") в номер (12)
                                    int condColIndex = ExcelColumnLetterToNumber(rule.ConditionColumn);

                                    if (condColIndex > 0)
                                    {
                                        // Читаем значение из ячейки условия
                                        string actualValue = ws.Cells[row, condColIndex].Text?.Trim();

                                        // Сравниваем с требуемым значением (ConditionValue)
                                        // Используем OrdinalIgnoreCase, чтобы "1" и "1 " или "да" и "ДА" совпадали
                                        bool isConditionMet = string.Equals(actualValue, rule.ConditionValue, StringComparison.OrdinalIgnoreCase);

                                        // ЕСЛИ УСЛОВИЕ НЕ ВЫПОЛНЕНО -> ПРОПУСКАЕМ ПРАВИЛО
                                        if (!isConditionMet)
                                        {
                                            continue;
                                        }
                                    }
                                }

                                // =========================================================
                                // ЭТАП 1: РАЗБОР СТРОКИ ПРАВИЛА (Search Logic)
                                // =========================================================
                                string effectiveSearchTerms = rule.ExcelValue;
                                string effectiveMasterName = rule.VisioMasterName;
                                string effectiveColumn = rule.SearchColumn;

                                // Логика "=" (Слово = Мастер ИЛИ Слово = Колонка = Мастер)
                                if (!string.IsNullOrEmpty(effectiveSearchTerms) && effectiveSearchTerms.Contains("="))
                                {
                                    var parts = effectiveSearchTerms.Split('=');
                                    if (parts.Length == 2)
                                    {
                                        effectiveSearchTerms = parts[0].Trim();
                                        effectiveMasterName = parts[1].Trim();
                                    }
                                    else if (parts.Length == 3)
                                    {
                                        effectiveSearchTerms = parts[0].Trim();
                                        effectiveColumn = parts[1].Trim();
                                        effectiveMasterName = parts[2].Trim();
                                    }
                                }

                                // =========================================================
                                // ЭТАП 2: ОПРЕДЕЛЕНИЕ ГДЕ ИСКАТЬ (Колонка или Весь ряд)
                                // =========================================================
                                string cellTextToSearch = "";

                                if (!string.IsNullOrWhiteSpace(effectiveColumn))
                                {
                                    // Если колонка задана (например "V"), берем текст строго оттуда
                                    int colIndex = ExcelColumnLetterToNumber(effectiveColumn);
                                    if (colIndex > 0)
                                    {
                                        cellTextToSearch = ws.Cells[row, colIndex].Text?.Trim();
                                    }
                                }
                                else
                                {
                                    // Если колонка НЕ задана, склеиваем первые 20 ячеек строки для поиска
                                    StringBuilder sb = new StringBuilder();
                                    for (int c = 1; c <= 20; c++) // Ограничиваемся 20 колонками для скорости
                                    {
                                        var txt = ws.Cells[row, c].Text?.Trim();
                                        if (!string.IsNullOrEmpty(txt)) sb.Append(txt + " ");
                                    }
                                    cellTextToSearch = sb.ToString().Trim();
                                }

                                if (string.IsNullOrEmpty(cellTextToSearch)) continue;

                                // =========================================================
                                // ЭТАП 3: ПОИСК СОВПАДЕНИЙ (Синонимы ; )
                                // =========================================================
                                var searchKeywords = effectiveSearchTerms.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                                var masterNames = effectiveMasterName.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                                int matchIndex = -1;
                                string foundRealName = "";

                                for (int i = 0; i < searchKeywords.Length; i++)
                                {
                                    string key = searchKeywords[i].Trim();
                                    if (string.IsNullOrEmpty(key)) continue;

                                    // Ищем вхождение (Contains)
                                    if (cellTextToSearch.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        matchIndex = i;
                                        foundRealName = cellTextToSearch; // Сохраняем полный текст для таблицы
                                        break;
                                    }
                                }

                                // =========================================================
                                // ЭТАП 4: ЕСЛИ НАШЛИ -> СОХРАНЯЕМ
                                // =========================================================
                                if (matchIndex >= 0)
                                {
                                    // Определяем мастера
                                    string targetMaster = "";
                                    if (masterNames.Length > 0)
                                    {
                                        if (matchIndex < masterNames.Length)
                                            targetMaster = masterNames[matchIndex].Trim();
                                        else
                                            targetMaster = masterNames[0].Trim();
                                    }

                                    // Определяем количество (Попытка найти цифру в колонке "Кол-во" или просто 1)
                                    int quantity = 1;
                                    // TODO: Если нужно читать кол-во из Excel, добавьте логику здесь.
                                    // Например: int qtyCol = ExcelColumnLetterToNumber("E"); 
                                    // int.TryParse(ws.Cells[row, qtyCol].Text, out quantity);

                                    // Добавляем результат
                                    rawHits.Add(new RawExcelHit
                                    {
                                        SheetName = sheetName,
                                        FullItemName = foundRealName,
                                        SearchTerm = searchKeywords[matchIndex],
                                        ConditionMet = true, // Мы это проверили на Этапе 0
                                        Quantity = quantity,
                                        IsLimited = rule.LimitQuantity,
                                        FoundRule = rule,
                                        TargetMasterName = targetMaster
                                    });

                                    // Прерываем цикл правил для этой строки, чтобы не дублировать
                                    break;
                                }
                            }
                        }
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

        private void ShowResultMessage(int totalFound)
        {
            if (totalFound > 0)
            {
                UpdateStatus($"✅ Сканирование завершено. Найдено {totalFound} позиций.");
            }
            else
            {
                UpdateStatus("⚠️ Сканирование завершено. Соответствующие позиции не найдены.");
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
            if (_rawHits == null || !_rawHits.Any())
            {
                MessageBox.Show("Нет данных! Сначала нажмите 'Загрузить Excel'.");
                return;
            }

            this.Enabled = false;
            UpdateStatus("⏳ Запуск Visio...");

            // Копируем данные для потока
            var hitsToProcess = _rawHits.ToList();

            await Task.Run(() =>
            {
                Visio.Application? visioApp = null;
                try
                {
                    visioApp = new Visio.Application();
                    visioApp.Visible = true;
                    var doc = visioApp.Documents.Add("");

                    // 1. Открываем ВСЕ трафареты заранее
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

                    // 2. Генерируем страницы (Прямой COM Interop)
                    GeneratePageDirectly(doc, "Маркировка", hitsToProcess, AppSettings.LabelingConfig);
                    GeneratePageDirectly(doc, "Схема", hitsToProcess, AppSettings.SchemeConfig);
                    GeneratePageDirectly(doc, "Шкаф", hitsToProcess, AppSettings.CabinetConfig);

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

        private void GeneratePageDirectly(Visio.Document doc, string pageName, List<RawExcelHit> hits, VisioConfiguration config)
        {
            Visio.Page page = null;
            try
            {
                try { page = doc.Pages.get_ItemU(pageName); }
                catch { page = doc.Pages.Add(); page.Name = pageName; }

                // Применяем настройки (размер, ориентация)
                SetupVisioPage(page, config);
            }
            catch { return; }

            // --- 1. ФИКСИРОВАННЫЕ ФИГУРЫ (Manual) ---
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
                        // Сначала бросаем в (0,0)
                        Visio.Shape shp = DropShapeOnPage(page, fixedItem.MasterName, 0, 0, 1, true);

                        if (shp != null)
                        {
                            string anchor = !string.IsNullOrWhiteSpace(fixedItem.Anchor) ? fixedItem.Anchor : "Center";
                            SetShapePosition(shp, x, y, anchor);
                        }
                    }
                }
            }

            // --- 2. НАЙДЕННЫЕ ФИГУРЫ (Sequential / Поток) ---
            if (config.SequentialDrawing.Enabled && hits != null && hits.Any())
            {
                double startX = 10, startY = 200;
                var sCoords = config.SequentialDrawing.StartCoordinatesXY?.Split(new[] { ',', ';' });
                if (sCoords != null && sCoords.Length >= 2)
                {
                    double.TryParse(sCoords[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out startX);
                    double.TryParse(sCoords[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out startY);
                }

                double curX = startX;
                double curY = startY;
                double maxW = config.SequentialDrawing.MaxLineWidthMM;
                double hGap = config.SequentialDrawing.HorizontalStepMM;
                double vGap = config.SequentialDrawing.VerticalStepMM;
                double rowMaxH = 0;

                // ГРУППИРОВКА ПО МАСТЕРУ
                // Используем TargetMasterName, чтобы разные слова могли вести к разным (или одинаковым) фигурам
                var groupedHits = hits
                    .Where(h => h.ConditionMet)
                    .GroupBy(h => new
                    {
                        Rule = h.FoundRule,
                        // Если TargetMasterName определен (новый код), используем его. Иначе fallback.
                        MasterName = !string.IsNullOrEmpty(h.TargetMasterName) ? h.TargetMasterName :
                                     (h.FoundRule != null ? h.FoundRule.VisioMasterName : h.SearchTerm)
                    });

                foreach (var group in groupedHits)
                {
                    var rule = group.Key.Rule;
                    string masterName = group.Key.MasterName;

                    if (string.IsNullOrWhiteSpace(masterName)) continue;

                    // Считаем количество
                    int countToDraw = (rule != null && rule.LimitQuantity) ? 1 : group.Sum(h => h.Quantity);

                    for (int i = 0; i < countToDraw; i++)
                    {
                        Visio.Shape shp = DropShapeOnPage(page, masterName, 0, 0, 1, true);

                        if (shp != null)
                        {
                            double wMM = shp.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters];
                            double hMM = shp.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                            // Перенос строки
                            if ((curX + wMM - startX) > maxW)
                            {
                                curX = startX;
                                curY -= (rowMaxH + vGap);
                                rowMaxH = 0;
                            }

                            string anchor = !string.IsNullOrWhiteSpace(rule.Anchor) ? rule.Anchor : config.SequentialDrawing.Anchor;
                            if (string.IsNullOrWhiteSpace(anchor)) anchor = "Center";

                            SetShapePosition(shp, curX, curY, anchor);

                            curX += wMM + hGap;
                            if (hMM > rowMaxH) rowMaxH = hMM;
                        }
                    }
                }
            }

            // --- 3. ОЧИСТКА ПУСТЫХ СТРАНИЦ ---
            // Удаляем стандартную "Страница-1", если она пустая и мы создали другие страницы
            try
            {
                if (doc.Pages.Count > 1)
                {
                    // Visio нумерует страницы с 1
                    var firstPage = doc.Pages[1];
                    // Проверка имен на разных языках
                    if ((firstPage.Name.StartsWith("Page") || firstPage.Name.StartsWith("Страница"))
                        && firstPage.Shapes.Count == 0)
                    {
                        firstPage.Delete(0);
                    }
                }
            }
            catch { /* Ошибка удаления игнорируется */ }
        }

        private void SetShapePosition(Visio.Shape shape, double xMM, double yMM, string anchor)
        {
            // 1. Получаем размеры в миллиметрах
            double w = shape.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters];
            double h = shape.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

            // --- ЛЕЧЕНИЕ ГЛЮКОВ (Нормализация фигуры) ---
            // Принудительно ставим Локальный Пин (точку вращения/привязки внутри фигуры) в её ЦЕНТР.
            // Это исправляет ситуацию, когда "Center ведет себя как TopLeft".
            if (shape.get_CellExists("LocPinX", 0) != 0)
                shape.get_CellsU("LocPinX").FormulaU = "Width*0.5";
            if (shape.get_CellExists("LocPinY", 0) != 0)
                shape.get_CellsU("LocPinY").FormulaU = "Height*0.5";
            // -------------------------------------------

            double finalX = xMM;
            double finalY = yMM;

            // Приводим к нижнему регистру для надежности
            switch (anchor?.ToLower().Replace("-", "").Trim())
            {
                case "topleft":
                    // Если x,y - это Левый-Верхний угол:
                    // Пин (Центр) должен быть правее на w/2 и ниже на h/2
                    finalX = xMM + (w / 2.0);
                    finalY = yMM - (h / 2.0);
                    break;

                case "bottomleft":
                    // Если x,y - это Левый-Нижний угол:
                    // Пин (Центр) должен быть правее на w/2 и выше на h/2
                    finalX = xMM + (w / 2.0);
                    finalY = yMM + (h / 2.0);
                    break;

                case "topright":
                    finalX = xMM - (w / 2.0);
                    finalY = yMM - (h / 2.0);
                    break;

                case "bottomright":
                    finalX = xMM - (w / 2.0);
                    finalY = yMM + (h / 2.0);
                    break;

                case "center":
                default:
                    // Если x,y - это Центр, и мы нормализовали LocPin выше,
                    // то координаты не меняем.
                    finalX = xMM;
                    finalY = yMM;
                    break;
            }

            // Применяем координаты (Visio всегда использует PinX/PinY для позиционирования на листе)
            shape.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = finalX;
            shape.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = finalY;
        }

        // Вспомогательный метод для вставки одной фигуры
        private Visio.Shape? DropShapeOnPage(Visio.Page page, string masterName, double xMM, double yMM, int qty, bool isSequential = false)
        {
            Visio.Master? mst = null;
            Visio.Document doc = page.Document;

            // 1. Ищем мастер в самом документе
            try { mst = doc.Masters.get_ItemU(masterName); } catch { }

            // 2. Если нет, ищем во всех открытых трафаретах
            if (mst == null)
            {
                foreach (Visio.Document d in doc.Application.Documents)
                {
                    if (d.Type == Visio.VisDocumentTypes.visTypeStencil)
                    {
                        try { mst = d.Masters.get_ItemU(masterName); } catch { }
                        if (mst != null) break;
                    }
                }
            }

            if (mst == null)
            {
                // Не нашли мастера — можно логировать ошибку
                return null;
            }

            Visio.Shape? lastShape = null;
            const double MM_TO_INCH = 1.0 / 25.4;

            for (int i = 0; i < qty; i++)
            {
                // Drop использует дюймы. PinX/PinY по умолчанию в центре фигуры
                // Если это Sequential, мы корректируем координаты в вызывающем коде, здесь просто ставим

                // Получаем размеры мастера чтобы скорректировать точку вставки (если Anchor TopLeft)
                // Но проще бросить, а потом подвинуть.

                lastShape = page.Drop(mst, xMM * MM_TO_INCH, yMM * MM_TO_INCH);

                // Для Sequential режим выравнивания обрабатывается в цикле выше
                // Для Manual можно добавить обработку Anchor здесь, если нужно
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
                UpdateStatus("⚠️ MasterMap пуст! Невозможно сопоставить данные.");
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
                    UpdateStatus($"  ✅ Сопоставлено: '{bestMatch}' -> '{visioMasterName}'");
                }

                if (!matched)
                {
                    UpdateStatus($"  ❌ Не сопоставлено: '{cleanedContent}'");
                }
            }

            UpdateStatus($"✅ Сопоставление завершено. Найдено {mappedItems} из {totalItems} позиций.");
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
                UpdateStatus("🔥 Начало генерации объединенного Visio-файла...");

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

                UpdateStatus($"✅ Файл Visio успешно сгенерирован и открыт");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                UpdateStatus($"❌ COM Ошибка Visio: {ex.Message}");
                MessageBox.Show($"COM Ошибка Visio: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                UpdateStatus($"❌ Общая ошибка при генерации Visio: {ex.Message}");
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
                            UpdateStatus($"❌ ОШИБКА: Файл трафарета не существует: {path}");
                            continue;
                        }

                        Visio.Document stencilDoc = page.Application.Documents.Open(path);
                        openStencils.Add(stencilDoc);
                        UpdateStatus($"✅ Трафарет открыт: {Path.GetFileName(path)}");
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"❌ КРИТИЧЕСКАЯ ОШИБКА открытия трафарета '{Path.GetFileName(path)}': {ex.Message}");
                    }
                }

                if (!openStencils.Any())
                {
                    UpdateStatus("❌ КРИТИЧЕСКАЯ ОШИБКА: Не удалось открыть ни один трафарет Visio. Прерывание.");
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

                            UpdateStatus($"  ✅ Добавлена предопределенная фигура: '{predefinedMasterName}'");
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"❌ Ошибка размещения предопределенного мастера '{predefinedMasterName}': {ex.Message}");
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
                        UpdateStatus($"❌ Мастер '{masterName}' отсутствует в списке доступных фигур.");
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
                        UpdateStatus($"❌ Не удалось получить COM-объект мастера '{masterName}', хотя он был найден в списке.");
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
                        UpdateStatus($"❌ Ошибка размещения мастера '{masterName}': {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObject(master);
                    }
                }

                UpdateStatus($"✅ Размещение фигур на странице '{page.Name}' завершено.");

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
                        "❌ КРИТИЧЕСКАЯ ОШИБКА: Не удалось разместить следующие фигуры (Мастера):\n\n" +
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
        public List<string> SelectedSheets { get; private set; } = new List<string>();

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



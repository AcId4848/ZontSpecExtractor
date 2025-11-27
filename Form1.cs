using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZontSpecExtractor.Properties;
using Visio = Microsoft.Office.Interop.Visio;


namespace ZontSpecExtractor
{
    // =========================================================================
    // 1. КОНФИГУРАЦИЯ И НАСТРОЙКИ (Сохранение состояния)
    // =========================================================================

    // Вне класса VisioConfiguration (но в том же namespace)
    public class PredefinedMasterConfig
    {
        public string MasterName { get; set; }
        public int Quantity { get; set; } = 1;
        public string CoordinatesXY { get; set; } // Пример: "10,20"
    }

    public class VisioConfiguration
    {
        public List<string> StencilFilePaths { get; set; } = new List<string>();

        // Основной список правил маппинга
        public List<SearchRule> SearchRules { get; set; } = new List<SearchRule>();

        // Список заранее определенных фигур
        public List<PredefinedMasterConfig> PredefinedMasterConfigs { get; set; } = new List<PredefinedMasterConfig>();

        [System.Text.Json.Serialization.JsonIgnore]
        public List<string> AvailableMasters { get; set; } = new List<string>();

        public string PageSize { get; set; } = "A4"; // Например: "A4", "A3"
        public string PageOrientation { get; set; } = "Portrait"; // Или "Landscape"

        public VisioConfiguration(bool isScheme)
        {
            if (isScheme)
            {
                SearchRules = new List<SearchRule>
                {
                    new SearchRule { ExcelValue = "H2000+proV2", VisioMasterName = "H2000+PRO ZONT Контроллер", SearchColumn = "C", UseCondition = true },
                    new SearchRule { ExcelValue = "H1500+pro", VisioMasterName = "H1500+PRO ZONT Контроллер", SearchColumn = "C", UseCondition = true }
                };
            }
            else
            {
                SearchRules = new List<SearchRule>
                {
                    new SearchRule { ExcelValue = "H2000+proV2", VisioMasterName = "Маркировка H2000 (пример)", SearchColumn = "C", UseCondition = true }
                };
            }
        }

        public VisioConfiguration() : this(true) { }
    }

    public class SearchRule
    {
        // Основное слово для поиска в Excel (используется и для поиска, и для маппинга Visio)
        public string ExcelValue { get; set; } = "";

        // === СВОЙСТВА ДЛЯ КАРТЫ VISIO (Excel=Col=Master) ===
        public string SearchColumn { get; set; } = ""; // ВОССТАНОВЛЕНО: Колонка, в которой ищется ExcelValue для маппинга

        // === НОВЫЕ СВОЙСТВА ДЛЯ ОБЩИХ НАСТРОЕК ПОИСКА ===
        // Если true, то даже если найдено 10 раз, в Visio попадет только 1 фигура
        public bool LimitQuantity { get; set; } = false;

        // Логика условий
        public bool UseCondition { get; set; } = false;

        // Значение условия (например "1", "V", "Да")
        public string ConditionValue { get; set; } = "";

        // Колонка, в которой проверяется условие (например "D" или "5")
        public string ConditionColumn { get; set; } = "";

        // Имя мастера Visio
        public string VisioMasterName { get; set; } = "";

        // Свойство для совместимости с кодом поиска
        [System.Text.Json.Serialization.JsonIgnore]
        public string Term { get => ExcelValue; set => ExcelValue = value; }
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
        private RichTextBox _rtxtSchemePaths, _rtxtLabelingPaths, _rtxtCabinetPaths;
        private RichTextBox _rtxtSchemeMap, _rtxtLabelingMap, _rtxtCabinetMap;
        private RichTextBox _rtxtSchemePredefined, _rtxtLabelingPredefined, _rtxtCabinetPredefined;
        private RichTextBox _rtxtSheetNames;
        private DataGridView _dgvSearchRules;

        // Вспомогательный метод для корректной очистки COM-объектов
        private static void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception) { /* Игнорировать ошибки при очистке */ }
            finally
            {
                obj = null;
            }
        }


        public GeneralSettingsForm()
        {
            this.Text = "Общие настройки";
            this.Size = new Size(1100, 850);
            this.StartPosition = FormStartPosition.CenterParent;
            SetupUI();
        }

        private void SetupUI()
        {
            var mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), RowCount = 2, RowStyles = { new RowStyle(SizeType.Percent, 100), new RowStyle(SizeType.Absolute, 50) } };
            var tabControl = new TabControl { Dock = DockStyle.Fill };

            var searchPage = new TabPage("Настройки поиска (Excel)") { Padding = new Padding(10) };
            searchPage.Controls.Add(CreateSearchConfigPanel());
            tabControl.TabPages.Add(searchPage);

            var visioPage = new TabPage("Настройки Visio") { Padding = new Padding(10) };
            visioPage.Controls.Add(CreateVisioConfigPanel());
            tabControl.TabPages.Add(visioPage);

            mainLayout.Controls.Add(tabControl, 0, 0);

            var btnSave = new Button { Text = "Сохранить", Width = 120, Height = 30, DialogResult = DialogResult.OK };
            btnSave.Click += BtnSave_Click;
            var btnCancel = new Button { Text = "Отмена", Width = 100, Height = 30, DialogResult = DialogResult.Cancel };

            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            footer.Controls.Add(btnCancel);
            footer.Controls.Add(btnSave);
            mainLayout.Controls.Add(footer, 0, 1);
            this.Controls.Add(mainLayout);
        }

        private Panel CreateSearchConfigPanel()
        {
            var panel = new Panel { Dock = DockStyle.Fill };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, RowStyles = { new RowStyle(SizeType.Absolute, 30), new RowStyle(SizeType.Absolute, 60), new RowStyle(SizeType.Absolute, 30), new RowStyle(SizeType.Percent, 100) } };

            layout.Controls.Add(new Label { Text = "Целевые листы:", Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);
            _rtxtSheetNames = new RichTextBox { Dock = DockStyle.Fill, Text = string.Join(Environment.NewLine, AppSettings.SearchConfig.TargetSheetNames), ReadOnly = true, BackColor = SystemColors.ControlLight };
            layout.Controls.Add(_rtxtSheetNames, 0, 1);

            layout.Controls.Add(new Label { Text = "Правила поиска:", Dock = DockStyle.Bottom, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 2);

            // --- ОБНОВЛЕНИЕ ТАБЛИЦЫ ---
            _dgvSearchRules = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, BackgroundColor = Color.White };

            // 1. Искомое слово
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Искомое слово",
                DataPropertyName = "ExcelValue",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            // 2. Ограничение (Галочка)
            _dgvSearchRules.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "Фикс. 1шт?",
                ToolTipText = "Если найдено много раз, добавлять только 1 раз?",
                DataPropertyName = "LimitQuantity",
                Width = 80
            });

            // 3. Условие (Галочка)
            _dgvSearchRules.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "Условие?",
                DataPropertyName = "UseCondition",
                Width = 70
            });

            // 4. Значение условия
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Значение условия",
                DataPropertyName = "ConditionValue",
                Width = 100
            });

            // 5. Колонка условия
            _dgvSearchRules.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Ячейка (Кол.)",
                ToolTipText = "Буква колонки (например D) где проверять условие",
                DataPropertyName = "ConditionColumn",
                Width = 90
            });

            // Visio Master (нужен для связки, хотя редактируется в другом месте, здесь можно оставить пустым или скрыть)
            // Но лучше оставить, чтобы понимать, что мы ищем, если маппинг идет отсюда. 
            // В вашей текущей логике маппинг идет в другом окне, но правила хранятся здесь.

            _dgvSearchRules.DataSource = new System.ComponentModel.BindingList<SearchRule>(AppSettings.SearchConfig.Rules ?? new List<SearchRule>());
            layout.Controls.Add(_dgvSearchRules, 0, 3);
            // ---------------------------

            panel.Controls.Add(layout);
            return panel;
        }

        private Panel CreateVisioConfigPanel()
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 2, RowStyles = { new RowStyle(SizeType.Absolute, 30), new RowStyle(SizeType.Percent, 100) }, ColumnStyles = { new ColumnStyle(SizeType.Percent, 33), new ColumnStyle(SizeType.Percent, 33), new ColumnStyle(SizeType.Percent, 33) } };
            layout.Controls.Add(new Label { Text = "Настройка соответствий (Excel=Col=Visio) и фигур (Name,Qty,X,Y)", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);
            layout.SetColumnSpan(layout.GetControlFromPosition(0, 0), 3);

            // Инициализируем RichTextBoxes, которые являются полями класса
            _rtxtLabelingPaths = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = SystemColors.ControlLight };
            _rtxtLabelingMap = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };
            _rtxtLabelingPredefined = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };

            _rtxtSchemePaths = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = SystemColors.ControlLight };
            _rtxtSchemeMap = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };
            _rtxtSchemePredefined = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };

            _rtxtCabinetPaths = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = SystemColors.ControlLight };
            _rtxtCabinetMap = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };
            _rtxtCabinetPredefined = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 8) };

            // Передаем их как обычные параметры
            layout.Controls.Add(CreateSingleConfig("1. МАРКИРОВКА", AppSettings.LabelingConfig, _rtxtLabelingPaths, _rtxtLabelingMap, _rtxtLabelingPredefined), 0, 1);
            layout.Controls.Add(CreateSingleConfig("2. СХЕМА", AppSettings.SchemeConfig, _rtxtSchemePaths, _rtxtSchemeMap, _rtxtSchemePredefined), 1, 1);
            layout.Controls.Add(CreateSingleConfig("3. ШКАФ", AppSettings.CabinetConfig, _rtxtCabinetPaths, _rtxtCabinetMap, _rtxtCabinetPredefined), 2, 1);

            return new Panel { Dock = DockStyle.Fill, Controls = { layout } };
        }

        private Panel CreateSingleConfig(string title, VisioConfiguration cfg, RichTextBox rPath, RichTextBox rMap, RichTextBox rPre)
        {
            var p = new Panel { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(2) };
            var l = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8, RowStyles = { new RowStyle(SizeType.Absolute, 20), new RowStyle(SizeType.Absolute, 20), new RowStyle(SizeType.Absolute, 60), new RowStyle(SizeType.Absolute, 30), new RowStyle(SizeType.Absolute, 20), new RowStyle(SizeType.Absolute, 80), new RowStyle(SizeType.Absolute, 20), new RowStyle(SizeType.Percent, 100) } };

            l.Controls.Add(new Label { Text = title, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 0);

            l.Controls.Add(new Label { Text = "Трафареты:", Dock = DockStyle.Bottom }, 0, 1);
            rPath.Text = string.Join(Environment.NewLine, cfg.StencilFilePaths); // Присваиваем текст
            l.Controls.Add(rPath, 0, 2);

            var btnP = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };

            // Кнопка "+" (Добавить трафареты)
            var btnAdd = new Button { Text = "+", Width = 30 };
            btnAdd.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Multiselect = true, Filter = "Visio|*.vssx;*.vsdx" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        // Добавляем только те, которых нет
                        cfg.StencilFilePaths.AddRange(ofd.FileNames.Where(f => !cfg.StencilFilePaths.Contains(f)));
                        rPath.Text = string.Join(Environment.NewLine, cfg.StencilFilePaths);
                    }
                }
            };
            btnP.Controls.Add(btnAdd);

            // Кнопка "X" (Удалить трафареты)
            var btnClear = new Button { Text = "X", Width = 30, ForeColor = Color.Red };
            btnClear.Click += (s, e) =>
            {
                if (MessageBox.Show("Очистить список трафаретов?", "Подтверждение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    cfg.StencilFilePaths.Clear();
                    rPath.Text = "";
                    cfg.AvailableMasters.Clear(); // Также очищаем список мастеров
                }
            };
            btnP.Controls.Add(btnClear);

            // КНОПКА ДЛЯ СКАНИРОВАНИЯ ФИГУР. ПЕРЕДАЕМ RichTextBox (rPre)
            var btnScan = new Button { Text = "Сканировать фигуры", AutoSize = true, BackColor = Color.LightGreen };
            btnScan.Click += (s, e) => ScanMasters(cfg, rPre);
            btnP.Controls.Add(btnScan);

            l.Controls.Add(btnP, 0, 3);

            l.Controls.Add(new Label { Text = "Фигуры (Name,Qty,X,Y):", Dock = DockStyle.Bottom, ForeColor = Color.Blue }, 0, 4);
            rPre.Text = string.Join(Environment.NewLine, cfg.PredefinedMasterConfigs.Select(c => $"{c.MasterName},{c.Quantity},{c.CoordinatesXY}"));
            l.Controls.Add(rPre, 0, 5);

            l.Controls.Add(new Label { Text = "Карта (Word=Col=Master):", Dock = DockStyle.Bottom }, 0, 6);
            rMap.Text = string.Join(Environment.NewLine, cfg.SearchRules.Select(r => $"{r.ExcelValue}={r.SearchColumn}={r.VisioMasterName}"));
            l.Controls.Add(rMap, 0, 7);

            p.Controls.Add(l);
            return p;
        }

        // Метод ShowMasterListWindow больше не нужен, так как фигуры сразу добавляются в RichTextBox

        // =========================================================================
        // ИСПРАВЛЕННЫЙ МЕТОД: ScanMasters (Теперь добавляет в RichTextBox)
        // =========================================================================
        private void ScanMasters(VisioConfiguration cfg, RichTextBox rPre)
        {
            if (!cfg.StencilFilePaths.Any())
            {
                MessageBox.Show("Сначала добавьте файлы трафаретов (*.vssx, *.vsdx)!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Visio.Application? visioApp = null;
            Visio.Document? stencilDoc = null;

            try
            {
                cfg.AvailableMasters.Clear();

                // 1. Запуск Visio
                visioApp = new Visio.Application();
                visioApp.Visible = false; // Работаем в фоновом режиме

                // 2. Проход по всем файлам трафаретов
                foreach (var path in cfg.StencilFilePaths.Where(File.Exists))
                {
                    // Открываем трафарет. visOpenHidden (4) для открытия без отображения
                    stencilDoc = visioApp.Documents.OpenEx(path, (short)Visio.VisOpenSaveArgs.visOpenHidden);

                    // Извлекаем все имена мастеров
                    foreach (Visio.Master master in stencilDoc.Masters)
                    {
                        if (!cfg.AvailableMasters.Contains(master.NameU))
                        {
                            cfg.AvailableMasters.Add(master.NameU);
                        }
                        ReleaseComObject(master);
                    }

                    // Закрываем трафарет
                    stencilDoc.Close();
                    ReleaseComObject(stencilDoc);
                    stencilDoc = null;
                }

                // 3. Сообщаем об успехе и предлагаем добавить фигуры
                string successMessage = $"Сканирование завершено.\nОбработано трафаретов: {cfg.StencilFilePaths.Count}\nНайдено уникальных фигур: {cfg.AvailableMasters.Count}";

                if (cfg.AvailableMasters.Any())
                {
                    // Копируем в буфер обмена для удобства
                    Clipboard.SetText(string.Join(Environment.NewLine, cfg.AvailableMasters.OrderBy(m => m)));
                    successMessage += "\n\nСписок фигур скопирован в буфер обмена.";

                    // НОВОЕ: Спрашиваем, нужно ли добавить фигуры в RichTextBox
                    var result = MessageBox.Show(
                        successMessage + "\n\nХотите добавить найденные фигуры в список предопределенных мастеров (Name,Qty,X,Y) в текущем окне?",
                        "Добавить фигуры",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (result == DialogResult.Yes)
                    {
                        // Форматируем новые мастера для вставки
                        var newMastersText = string.Join(Environment.NewLine, cfg.AvailableMasters
                            .OrderBy(m => m)
                            .Select(m => $"{m},1,0,0")); // Формат: MasterName,Quantity,X,Y

                        // Добавляем к текущему содержимому RichTextBox
                        string currentText = rPre.Text.Trim();
                        if (!string.IsNullOrEmpty(currentText))
                        {
                            // Добавляем новые элементы, проверяя, чтобы избежать дубликатов в RichTextBox 
                            // (хотя дубликаты могут возникнуть, если пользователь ранее вручную ввел)
                            rPre.Text = currentText + Environment.NewLine + newMastersText;
                        }
                        else
                        {
                            rPre.Text = newMastersText;
                        }

                        MessageBox.Show("Фигуры добавлены в текстовое поле 'Фигуры (Name,Qty,X,Y)'.\n\nНе забудьте нажать 'Сохранить'!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show(successMessage, "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сканирования Visio: {ex.Message}", "Ошибка Visio COM", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 4. Очистка COM-объектов
                if (stencilDoc != null)
                {
                    try { stencilDoc.Close(); } catch { }
                    ReleaseComObject(stencilDoc);
                }
                if (visioApp != null)
                {
                    try { visioApp.Quit(); } catch { }
                    ReleaseComObject(visioApp);
                }
                // Вызов сборщика мусора для полной очистки COM-ссылок
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // ... (остальные методы ParseRules, ParsePre, ParsePredefinedConfigs, ParseList, ParseSearchRules, BtnSave_Click) ...

        private List<SearchRule> ParseRules(string text)
        {
            var res = new List<SearchRule>();
            foreach (var line in text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
            {
                var p = line.Split('=').Select(x => x.Trim()).ToArray();
                if (p.Length >= 3) res.Add(new SearchRule { ExcelValue = p[0], SearchColumn = p[1], VisioMasterName = p[2] });
                else if (p.Length == 2) res.Add(new SearchRule { ExcelValue = p[0], SearchColumn = "C", VisioMasterName = p[1] });
            }
            return res;
        }

        private List<PredefinedMasterConfig> ParsePre(string text)
        {
            var res = new List<PredefinedMasterConfig>();
            foreach (var line in text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
            {
                var p = line.Split(',').Select(x => x.Trim()).ToArray();
                if (p.Length > 0)
                {
                    var c = new PredefinedMasterConfig { MasterName = p[0] };
                    if (p.Length > 1 && int.TryParse(p[1], out int q)) c.Quantity = q;
                    if (p.Length >= 3) c.CoordinatesXY = $"{p[2]},{(p.Length > 3 ? p[3] : "0")}";
                    else c.CoordinatesXY = "0,0";
                    res.Add(c);
                }
            }
            return res;
        }

        private List<string> ParseList(RichTextBox rtb)
        {
            return rtb.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim())
                            .ToList();
        }

        private List<SearchRule> ParseSearchRules(RichTextBox rtb)
        {
            var rules = new List<SearchRule>();
            var lines = rtb.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var parts = line.Split('=').Select(p => p.Trim()).ToArray();
                if (parts.Length == 3)
                {
                    rules.Add(new SearchRule
                    {
                        ExcelValue = parts[0],
                        SearchColumn = parts[1],
                        VisioMasterName = parts[2]
                    });
                }
            }
            return rules;
        }

        private List<PredefinedMasterConfig> ParsePredefinedConfigs(RichTextBox rtb)
        {
            var configs = new List<PredefinedMasterConfig>();
            var lines = rtb.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var parts = line.Split(',').Select(p => p.Trim()).ToArray();
                if (parts.Length >= 1) // MasterName, [Quantity], [X], [Y]
                {
                    var masterName = parts[0];
                    int quantity = parts.Length > 1 && int.TryParse(parts[1], out int qty) ? qty : 1;
                    string coords = "0,0";

                    if (parts.Length >= 3)
                    {
                        if (parts.Length >= 4) coords = $"{parts[2]},{parts[3]}";
                        else coords = $"{parts[2]},0";
                    }

                    if (!string.IsNullOrEmpty(masterName))
                    {
                        configs.Add(new PredefinedMasterConfig
                        {
                            MasterName = masterName,
                            Quantity = quantity,
                            CoordinatesXY = coords
                        });
                    }
                }
            }
            return configs;
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            try
            {
                // 1. Обновляем настройки Visio (SearchRules вместо MasterMap)
                AppSettings.SchemeConfig.SearchRules = ParseSearchRules(_rtxtSchemeMap);
                AppSettings.LabelingConfig.SearchRules = ParseSearchRules(_rtxtLabelingMap);
                AppSettings.CabinetConfig.SearchRules = ParseSearchRules(_rtxtCabinetMap);

                // 2. Обновляем предопределенные мастера
                AppSettings.SchemeConfig.PredefinedMasterConfigs = ParsePredefinedConfigs(_rtxtSchemePredefined);
                AppSettings.LabelingConfig.PredefinedMasterConfigs = ParsePredefinedConfigs(_rtxtLabelingPredefined);
                AppSettings.CabinetConfig.PredefinedMasterConfigs = ParsePredefinedConfigs(_rtxtCabinetPredefined);

                // 3. Обновляем целевые листы для поиска
                AppSettings.SearchConfig.TargetSheetNames = ParseList(_rtxtSheetNames);

                // --- СОХРАНЕНИЕ ПРАВИЛ ПОИСКА ИЗ ТАБЛИЦЫ ---
                var newRules = new List<SearchRule>();
                if (_dgvSearchRules.DataSource is System.ComponentModel.BindingList<SearchRule> list)
                {
                    newRules = list.Where(r => !string.IsNullOrWhiteSpace(r.Term)).ToList();
                }
                AppSettings.SearchConfig.Rules = newRules;
                // -------------------------------------------

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

        private async void LoadFiles(string[] filePaths)
        {
            dataGridView.Rows.Clear();
            data.Clear();

            if (filePaths.Length > 0)
            {
                lblFileInfo.Text = $"📄 {System.IO.Path.GetFileName(filePaths[0])}";
            }
            else
            {
                lblFileInfo.Text = "Файл не выбран";
                return;
            }

            int totalFound = 0;
            UpdateStatus("⚙️ Идет сканирование Excel-файла...");

            await Task.Run(() =>
            {
                foreach (var path in filePaths)
                {
                    var rows = ScanSpecificSheet(path);
                    data.AddRange(rows);
                    totalFound += rows.Count;
                }
            });

            UpdateDataGridView();
            ShowResultMessage(totalFound);
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
        private List<Dictionary<string, string>> ScanSpecificSheet(string path)
        {
            var finalRows = new List<Dictionary<string, string>>();
            var targetSheets = AppSettings.SearchConfig.TargetSheetNames;
            var rules = AppSettings.SearchConfig.Rules;

            if (rules == null || !rules.Any()) return finalRows;

            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                foreach (string sheetName in targetSheets)
                {
                    var ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                    if (ws == null || ws.Dimension == null) continue;

                    int startRow = ws.Dimension.Start.Row;
                    int endRow = ws.Dimension.End.Row;
                    int endCol = ws.Dimension.End.Column;

                    // Кэш значений для правил "Ограниченное кол-во"
                    // Словарь: ИмяПравила -> Найдены ли уже совпадения?
                    var limitedRulesFound = new HashSet<string>();

                    // Сканируем построчно
                    for (int row = startRow; row <= endRow; row++)
                    {
                        // Получаем содержимое строки (кэшируем первые 30 колонок для скорости поиска слова)
                        var rowValues = new List<string>();
                        int maxColToScan = Math.Min(endCol, 30);
                        for (int c = 1; c <= maxColToScan; c++)
                        {
                            rowValues.Add(ws.Cells[row, c].Text.Trim());
                        }

                        foreach (var rule in rules)
                        {
                            if (string.IsNullOrWhiteSpace(rule.Term)) continue;

                            bool match = false;
                            string foundQty = "0";

                            // 1. Ищем слово (Term) в строке
                            if (rowValues.Any(v => v.IndexOf(rule.Term, StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                // 2. Проверка условия
                                if (rule.UseCondition)
                                {
                                    // Получаем индекс колонки условия (например "D" -> 4)
                                    int conditionColIndex = GetColumnIndex(rule.ConditionColumn);

                                    if (conditionColIndex > 0 && conditionColIndex <= endCol)
                                    {
                                        string cellValue = ws.Cells[row, conditionColIndex].Text.Trim();
                                        // Проверяем совпадение значения
                                        if (cellValue.Equals(rule.ConditionValue, StringComparison.OrdinalIgnoreCase))
                                        {
                                            match = true;
                                        }
                                    }
                                }
                                else
                                {
                                    // Условие не используется - просто нашли слово
                                    match = true;
                                }
                            }

                            if (match)
                            {
                                // Определяем количество
                                if (rule.LimitQuantity)
                                {
                                    // Логика "Ограниченное количество":
                                    // Мы записываем "1", но при группировке мы учтем это правило.
                                    // Или можно просто добавлять, а на этапе группировки срезать.
                                    foundQty = "1";
                                }
                                else
                                {
                                    // Обычная логика: ищем число в строке (эвристика) или берем 1
                                    var numStr = rowValues.FirstOrDefault(v =>
                                        double.TryParse(v, out double d) && d > 0 && d < 1000);
                                    foundQty = numStr ?? "1";
                                }

                                // Добавляем запись. 
                                // Важно: сохраняем ссылку на правило (LimitQuantity), чтобы использовать при группировке.
                                finalRows.Add(new Dictionary<string, string>
                                {
                                    ["Лист"] = sheetName,
                                    ["Наименование"] = rule.Term,
                                    ["Количество"] = foundQty,
                                    ["_IsLimited"] = rule.LimitQuantity.ToString() // Внутренний флаг
                                });

                                // Прерываем цикл правил для этой строки (чтобы одну строку не посчитать дважды для разных правил?)
                                // Зависит от задачи. Обычно лучше break.
                                break;
                            }
                        }
                    }
                }
            }

            // ГРУППИРОВКА И ПОДСЧЕТ ИТОГОВ
            return finalRows
                .GroupBy(r => new { Sheet = r["Лист"], Name = r["Наименование"] })
                .Select(g =>
                {
                    // Проверяем, было ли у этого правила ограничение (смотрим на флаг первой попавшейся записи группы)
                    bool isLimited = bool.TryParse(g.First().GetValueOrDefault("_IsLimited"), out bool limit) && limit;

                    int totalQty;
                    if (isLimited)
                    {
                        // Если стоит галочка "Ограничить", то независимо от количества найденных строк, сумма = 1
                        totalQty = 1;
                    }
                    else
                    {
                        // Иначе суммируем всё, что нашли
                        totalQty = g.Sum(x => int.TryParse(x["Количество"], out int q) ? q : 0);
                    }

                    return new Dictionary<string, string>
                    {
                        ["Лист"] = g.Key.Sheet,
                        ["Наименование"] = g.Key.Name,
                        ["Количество"] = totalQty.ToString()
                    };
                })
                .Where(x => x["Количество"] != "0")
                .ToList();
        }

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

        private void OpenVisioSettingsClick(object? sender, EventArgs e)
        {
            using (var settingsForm = new GeneralSettingsForm())
            {
                if (settingsForm.ShowDialog(this) == DialogResult.OK)
                {
                    UpdateStatus("Общие настройки сохранены.");
                }
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
        private static void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception) { /* Игнорировать ошибки при очистке */ }
        }

        public static void RunDrawingMacro(Visio.Document doc, Visio.Page page, List<VisioItem> itemsToDraw, VisioConfiguration config)
        {
            if (itemsToDraw == null || itemsToDraw.Count == 0) return;

            string pageName = page.Name;
            string moduleName = "Mod_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            StringBuilder sb = new StringBuilder();

            var form1Instance = Form.ActiveForm as Form1;
            if (form1Instance == null) return;

            var (width, height) = form1Instance.GetPageDimensions(config.PageSize, config.PageOrientation);
            string pageWidth = width.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string pageHeight = height.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // --- НАЧАЛО VBA КОДА ---
            sb.AppendLine($"Sub Draw_{moduleName}()");
            sb.AppendLine($"    Dim pg As Visio.Page");
            sb.AppendLine($"    Set pg = ActiveDocument.Pages.ItemU(\"{pageName}\")");
            sb.AppendLine($"    Dim mst As Visio.Master");
            sb.AppendLine($"    Dim doc As Visio.Document");
            sb.AppendLine($"    Dim i As Integer");
            sb.AppendLine($"    Dim found As Boolean");

            // --- ДАННЫЕ: БЕЗОПАСНОЕ СОЗДАНИЕ МАССИВОВ ---
            int count = itemsToDraw.Count;
            int maxIndex = count - 1;

            sb.AppendLine($"    ReDim mastersArr(0 To {maxIndex}) As String");
            sb.AppendLine("");

            for (int k = 0; k < count; k++)
            {
                var item = itemsToDraw[k];

                string safeName = item.MasterName
                    .Replace("\r", "")
                    .Replace("\n", "")
                    .Replace("\"", "\"\"")
                    .Trim();

                sb.AppendLine($"    mastersArr({k}) = \"{safeName}\"");
            }
            sb.AppendLine("    ' --------------------------------------------------");


            // --- ЛОГИКА АВТОМАТИЧЕСКОГО РАЗМЕЩЕНИЯ (PACKING) ---
            sb.AppendLine("");
            sb.AppendLine($"    ' --- КОНСТАНТЫ РАЗМЕРА ЛИСТА (В ДЮЙМАХ) ---");
            sb.AppendLine($"    Const PAGE_WIDTH As Double = {pageWidth}");
            sb.AppendLine($"    Const PAGE_HEIGHT As Double = {pageHeight}");
            sb.AppendLine($"    Const MARGIN As Double = 0.1");
            sb.AppendLine($"    Dim CurrentX As Double");
            sb.AppendLine($"    Dim CurrentY As Double");
            sb.AppendLine($"    Dim MaxRowHeight As Double: MaxRowHeight = 0");

            sb.AppendLine($"    CurrentY = PAGE_HEIGHT - MARGIN");
            sb.AppendLine($"    CurrentX = MARGIN");
            sb.AppendLine("");

            sb.AppendLine("    For i = LBound(mastersArr) To UBound(mastersArr)");
            sb.AppendLine("        Set mst = Nothing");
            sb.AppendLine("        found = False");
            sb.AppendLine("        Dim mName As String");
            sb.AppendLine("        mName = mastersArr(i)");

            // 1. Сначала ищем в самом документе (вдруг уже есть)
            sb.AppendLine("        On Error Resume Next");
            sb.AppendLine("        Set mst = ActiveDocument.Masters.ItemU(mName)");
            sb.AppendLine("        On Error GoTo 0");

            // 2. Если не нашли, ищем во ВСЕХ открытых трафаретах
            sb.AppendLine("        If mst Is Nothing Then");
            sb.AppendLine("            For Each doc In Application.Documents");
            sb.AppendLine("                If doc.Type = 2 Then");
            sb.AppendLine("                    On Error Resume Next");
            sb.AppendLine("                    Set mst = doc.Masters.ItemU(mName)");
            sb.AppendLine("                    On Error GoTo 0");
            sb.AppendLine("                    If Not mst Is Nothing Then Exit For");
            sb.AppendLine("                End If");
            sb.AppendLine("            Next doc");
            sb.AppendLine("        End If");

            // ----------------------------------------------------------------------
            // DEBUG CODE: Проверка, был ли найден мастер
            // ----------------------------------------------------------------------
            sb.AppendLine("        If mst Is Nothing Then");
            sb.AppendLine("            Debug.Print \"Мастер НЕ НАЙДЕН: \" & mName");
            sb.AppendLine("        End If");
            // ----------------------------------------------------------------------

            // 3. Если нашли мастера - рисуем
            sb.AppendLine("        If Not mst Is Nothing Then");
            sb.AppendLine("            ' 1. Получаем размеры мастера (Ширина/Высота в дюймах)");
            sb.AppendLine("            Dim w As Double: w = mst.Cells(\"Width\").Result(visInches)");
            sb.AppendLine("            Dim h As Double: h = mst.Cells(\"Height\").Result(visInches)");
            sb.AppendLine("            ");
            sb.AppendLine("            ' 2. Проверяем, помещается ли фигура в текущей строке:");
            sb.AppendLine("            If (CurrentX + w + MARGIN) > PAGE_WIDTH Then");
            sb.AppendLine("                ' Переход на новую строку");
            sb.AppendLine("                CurrentY = CurrentY - MaxRowHeight - MARGIN");
            sb.AppendLine("                CurrentX = MARGIN");
            sb.AppendLine("                MaxRowHeight = 0");
            sb.AppendLine("            End If");
            sb.AppendLine("            ");
            sb.AppendLine("            ' 3. Проверяем, не вышли ли мы за нижний край страницы");
            sb.AppendLine("            If (CurrentY - h - MARGIN) < 0 Then");

            // ----------------------------------------------------------------------
            // DEBUG CODE: Сообщение о выходе из-за нехватки места
            // ----------------------------------------------------------------------
            sb.AppendLine("                MsgBox \"РИСОВАНИЕ ОСТАНОВЛЕНО: Не хватило места на странице для: \" & mName, vbCritical");
            // ----------------------------------------------------------------------

            sb.AppendLine("                Exit For");
            sb.AppendLine("            End If");
            sb.AppendLine("            ");
            sb.AppendLine("            ' 4. Обновляем максимальную высоту для текущей строки");
            sb.AppendLine("            If h > MaxRowHeight Then MaxRowHeight = h");
            sb.AppendLine("            ");
            sb.AppendLine("            ' 5. Вычисляем координаты центра фигуры (для метода Drop)");
            sb.AppendLine("            Dim CenterX As Double: CenterX = CurrentX + w / 2");
            sb.AppendLine("            Dim CenterY As Double: CenterY = CurrentY - h / 2");
            sb.AppendLine("            ");
            sb.AppendLine("            ' 6. Размещаем фигуру");
            sb.AppendLine("            pg.Drop mst, CenterX, CenterY");
            sb.AppendLine("            ");
            sb.AppendLine("            ' 7. Обновляем CurrentX для следующей фигуры");
            sb.AppendLine("            CurrentX = CurrentX + w + MARGIN");
            sb.AppendLine("        End If");

            sb.AppendLine("    Next i");
            sb.AppendLine($"End Sub");
            // --- КОНЕЦ VBA КОДА ---

            // --- БЕЗОПАСНЫЙ ЗАПУСК И ОЧИСТКА ---
            Microsoft.Vbe.Interop.VBProject vbProject = null;
            Microsoft.Vbe.Interop.VBComponent vbComp = null;

            try
            {
                vbProject = doc.VBProject;
                vbComp = vbProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                vbComp.Name = moduleName;

                vbComp.CodeModule.AddFromString(sb.ToString());

                doc.ExecuteLine($"Draw_{moduleName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Ошибка запуска VBA. Проверьте настройки безопасности Visio (доверять доступ к объектной модели проекта VBA).\n\n" +
                    ex.Message,
                    "Ошибка VBA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (vbComp != null)
                {
                    try { vbProject.VBComponents.Remove(vbComp); } catch { }
                    ReleaseComObject(vbComp);
                }
                if (vbProject != null) ReleaseComObject(vbProject);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

    // Вспомогательный класс для передачи данных
    public class VisioItem
    {
        public string MasterName { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
    }

        private async void OpenVisioClick(object? sender, EventArgs e)
        {
            var extractedData = this.data;
            if (extractedData == null || extractedData.Count == 0)
            {
                MessageBox.Show("Нет данных для отрисовки!");
                return;
            }

            this.Enabled = false;
            UpdateStatus("⏳ Идет генерация Visio через VBA...");

            await Task.Run(() =>
            {
                try
                {
                    var visioApp = new Visio.Application();
                    visioApp.Visible = true;
                    var doc = visioApp.Documents.Add(""); // Создаем новый документ

                    // --- ЛОГИКА ИМЕНОВАНИЯ И СОХРАНЕНИЯ ---
                    // Берем имя первого найденного листа из данных или из настроек
                    string sheetName = extractedData.FirstOrDefault()?["Лист"]
                                       ?? AppSettings.SearchConfig.TargetSheetNames.FirstOrDefault()
                                       ?? "ZONT";

                    // Очищаем имя от недопустимых символов файла
                    foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                    {
                        sheetName = sheetName.Replace(c, '_');
                    }

                    string suggestedName = $"Схемы_{sheetName}.vsdx";

                    // Чтобы сохранить файл, нам нужно вернуться в UI поток для диалога
                    // Или сохранить во временную папку.
                    // Лучший вариант для UX - спросить пользователя сразу или сохранить в "Мои документы"

                    string tempPath = System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), suggestedName);

                    // Сохраняем документ (Visio требует полного пути)
                    // Используем Invoke, если хотим спросить пользователя, но здесь сохраним автоматически для скорости
                    try
                    {
                        doc.SaveAs(tempPath);
                        UpdateStatus($"📁 Файл сохранен как: {tempPath}");
                    }
                    catch (Exception saveEx)
                    {
                        // Если файл занят или ошибка, продолжаем без сохранения имени
                        UpdateStatus($"⚠️ Не удалось сохранить имя файла автоматически: {saveEx.Message}");
                    }
                    // ---------------------------------------

                    // 1. Создаем/настраиваем страницы
                    Visio.Page pMarking = doc.Pages[1]; pMarking.Name = "Маркировка";
                    Visio.Page pScheme = doc.Pages.Add(); pScheme.Name = "Схема";
                    Visio.Page pCabinet = doc.Pages.Add(); pCabinet.Name = "Шкаф";

                    // 2. Открываем трафареты
                    var allPaths = AppSettings.LabelingConfig.StencilFilePaths
                        .Union(AppSettings.SchemeConfig.StencilFilePaths)
                        .Union(AppSettings.CabinetConfig.StencilFilePaths)
                        .Distinct().ToList();

                    foreach (var path in allPaths)
                    {
                        if (File.Exists(path)) visioApp.Documents.OpenEx(path, (short)Visio.VisOpenSaveArgs.visOpenDocked);
                    }

                    // 3. Запускаем генерацию
                    ProcessPageVba(doc, pMarking, extractedData, AppSettings.LabelingConfig);
                    ProcessPageVba(doc, pScheme, extractedData, AppSettings.SchemeConfig);
                    ProcessPageVba(doc, pCabinet, extractedData, AppSettings.CabinetConfig);

                    // В конце снова сохраняем, чтобы записать изменения
                    try { doc.Save(); } catch { }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
            });

            this.Enabled = true;
            UpdateStatus("✅ Готово.");
        }

        // ВСПОМОГАТЕЛЬНЫЙ МЕТОД ДЛЯ СБОРКИ ДАННЫХ
        private void ProcessPageVba(Visio.Document doc, Visio.Page page,
            List<Dictionary<string, string>> data, VisioConfiguration config)
        {
            var masterMap = RulesToMap(config.SearchRules);
            var itemsToDraw = new List<VisioItem>();
            double curX = 1.0, curY = 10.0; // Начало координат

            // А. Предопределенные фигуры (без Excel)
            foreach (var pm in config.PredefinedMasterConfigs)
            {
                if (pm == null || string.IsNullOrWhiteSpace(pm.MasterName))
                    continue;

                itemsToDraw.Add(new VisioItem
                {
                    MasterName = pm.MasterName.Trim(),
                    X = curX,
                    Y = curY
                });

                curX += 1.5;
                if (curX > 8) { curX = 1.0; curY -= 1.5; }
            }

            // Б. Фигуры из Excel
            // Используем существующий PrepareVisioData или ищем вручную
            // Здесь упрощенная логика поиска по карте:
            foreach (var row in data)
            {
                string name = row["Наименование"];
                // Ищем, есть ли такое имя (или его часть) в MasterMap ключах
                var matchedKey = config.SearchRules
                    .Select(r => r.ExcelValue)
                    .FirstOrDefault(k => name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);


                if (matchedKey != null)
                {
                    string masterName = config.SearchRules
                        .First(r => r.ExcelValue == matchedKey)
                        .VisioMasterName;

                    int qty = int.Parse(row.GetValueOrDefault("Количество", "1"));

                    for (int i = 0; i < qty; i++)
                    {
                        itemsToDraw.Add(new VisioItem { MasterName = masterName, X = curX, Y = curY });
                        curX += 1.5;
                        if (curX > 8) { curX = 1.0; curY -= 1.5; }
                    }
                }
            }

            // В. Запуск макроса
            VisioVbaRunner.RunDrawingMacro(doc, page, itemsToDraw, AppSettings.SchemeConfig);
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
            const double SPACING = 0.05; // 5 см между фигурами
            const double PAGE_WIDTH = 0.297; // A4 width in meters (29.7cm)

            double currentX = 1.0;
            double currentY = 1.0;

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
                        master = stencilDoc.Masters.Cast<Visio.Master>().FirstOrDefault(m =>
                            m.NameU.Equals(predefinedMasterName, StringComparison.Ordinal));

                        if (master != null) break;
                    }

                    if (master != null)
                    {
                        try
                        {
                            // 3.2. Добавление фигуры
                            // Здесь можно использовать координаты из конфига, если они есть, но пока оставим вашу логику currentX/Y
                            shape = page.Drop(master, currentX, currentY);
                            shape.Text = $"Предопределенная: {predefinedMasterName}";

                            // Логика смещения
                            currentX += 1.0;
                            if (currentX > 10.0)
                            {
                                currentX = 1.0;
                                currentY += 1.0;
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
                            shape = page.Drop(master, currentX, currentY);
                            shape.Text = masterName;

                            // Логика смещения
                            currentX += 1.0;
                            if (currentX > 10.0)
                            {
                                currentX = 1.0;
                                currentY += 1.0;
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
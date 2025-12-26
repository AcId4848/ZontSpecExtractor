using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZontSpecExtractor
{
    public partial class DebugPanelForm : Form
    {
        private RichTextBox _logTextBox;
        private TextBox _searchTextBox;
        private CheckBox _autoSaveCheckBox;
        private NumericUpDown _autoSaveInterval;
        private Label _memoryLabel;
        private Label _cpuLabel;
        private System.Windows.Forms.Timer _metricsTimer;
        private bool _colorMode = true;
        private bool _isPaused = false;
        private string _lastSearchTerm = "";
        private List<LogEntry> _allLogs = new List<LogEntry>();
        private readonly object _logsLock = new object();

        public DebugPanelForm()
        {
            InitializeComponent();
            SetupEventHandlers();
            StartMetricsTimer();
        }

        private void InitializeComponent()
        {
            this.Text = "🔴 Панель управления отладкой - КРАСНАЯ КНОПКА";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(800, 600);

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(5)
            };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 90));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 10));

            // === ЛЕВАЯ ПАНЕЛЬ: ЛОГ ===
            var logPanel = new Panel { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle };
            _logTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9),
                ReadOnly = true,
                BackColor = Color.Black,
                ForeColor = Color.LightGreen
            };
            logPanel.Controls.Add(_logTextBox);
            mainLayout.Controls.Add(logPanel, 0, 0);

            // === ПРАВАЯ ПАНЕЛЬ: КНОПКИ ===
            var buttonsPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            var buttonsFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(5)
            };

            // 1. 🔴 КНОПКА ПАНИКИ
            var btnPanic = CreateButton("🔴 КНОПКА ПАНИКИ", Color.Red, Color.White);
            btnPanic.Click += BtnPanic_Click;
            buttonsFlow.Controls.Add(btnPanic);

            // 2. 📤 Отправить лог в Telegram
            var btnSendTG = CreateButton("📤 Отправить лог в TG", Color.Orange, Color.White);
            btnSendTG.Click += BtnSendTG_Click;
            buttonsFlow.Controls.Add(btnSendTG);

            // 3. 💾 Сохранить на ПК
            var btnSave = CreateButton("💾 Сохранить на ПК", Color.Blue, Color.White);
            btnSave.Click += BtnSave_Click;
            buttonsFlow.Controls.Add(btnSave);

            // 3.5. 📁 Открыть папку с логами
            var btnOpenLogsFolder = CreateButton("📁 Открыть папку логов", Color.DarkBlue, Color.White);
            btnOpenLogsFolder.Click += BtnOpenLogsFolder_Click;
            btnOpenLogsFolder.Height = 30; // Делаем кнопку немного меньше
            buttonsFlow.Controls.Add(btnOpenLogsFolder);

            // 4. 🔄 Автосохранение
            var autoSavePanel = new Panel { Height = 60, Width = 200 };
            _autoSaveCheckBox = new CheckBox { Text = "🔄 Автосохранение", AutoSize = true, Location = new Point(5, 5) };
            _autoSaveInterval = new NumericUpDown { Minimum = 10, Maximum = 3600, Value = 60, Width = 80, Location = new Point(5, 30) };
            var lblInterval = new Label { Text = "сек", AutoSize = true, Location = new Point(90, 32) };
            _autoSaveCheckBox.CheckedChanged += AutoSaveCheckBox_CheckedChanged;
            autoSavePanel.Controls.AddRange(new Control[] { _autoSaveCheckBox, _autoSaveInterval, lblInterval });
            buttonsFlow.Controls.Add(autoSavePanel);

            // 5. 🧹 Очистить консоль
            var btnClear = CreateButton("🧹 Очистить консоль", Color.Gray, Color.White);
            btnClear.Click += BtnClear_Click;
            buttonsFlow.Controls.Add(btnClear);

            // 6. 🔍 Поиск/Фильтр
            var searchPanel = new Panel { Height = 50, Width = 200 };
            var lblSearch = new Label { Text = "🔍 Поиск:", AutoSize = true, Location = new Point(5, 5) };
            _searchTextBox = new TextBox { Width = 190, Location = new Point(5, 25) };
            _searchTextBox.TextChanged += SearchTextBox_TextChanged;
            searchPanel.Controls.AddRange(new Control[] { lblSearch, _searchTextBox });
            buttonsFlow.Controls.Add(searchPanel);

            // 7. ⏸ Приостановить логирование
            var btnPause = CreateButton("⏸ Приостановить логирование", Color.Yellow, Color.Black);
            btnPause.Click += BtnPause_Click;
            buttonsFlow.Controls.Add(btnPause);

            // 8. 📋 Копировать в буфер
            var btnCopy = CreateButton("📋 Копировать в буфер", Color.Purple, Color.White);
            btnCopy.Click += BtnCopy_Click;
            buttonsFlow.Controls.Add(btnCopy);

            // 9. 📉 Использование памяти
            var metricsPanel = new Panel { Height = 60, Width = 200 };
            _memoryLabel = new Label { Text = "Память: N/A", AutoSize = true, Location = new Point(5, 5), ForeColor = Color.Cyan };
            _cpuLabel = new Label { Text = "CPU: N/A", AutoSize = true, Location = new Point(5, 25), ForeColor = Color.Cyan };
            metricsPanel.Controls.AddRange(new Control[] { _memoryLabel, _cpuLabel });
            buttonsFlow.Controls.Add(metricsPanel);

            // 10. 📧 Email отчет (Заглушка)
            var btnEmail = CreateButton("📧 Email отчет", Color.Teal, Color.White);
            btnEmail.Click += BtnEmail_Click;
            buttonsFlow.Controls.Add(btnEmail);

            // 11. 🎨 Цветовой режим
            var btnColor = CreateButton("🎨 Цветовой режим", Color.Magenta, Color.White);
            btnColor.Click += BtnColor_Click;
            buttonsFlow.Controls.Add(btnColor);

            // 12. 🐛 Внедрить тестовую ошибку
            var btnTestError = CreateButton("🐛 Внедрить тестовую ошибку", Color.DarkRed, Color.White);
            btnTestError.Click += BtnTestError_Click;
            buttonsFlow.Controls.Add(btnTestError);

            buttonsPanel.Controls.Add(buttonsFlow);
            mainLayout.Controls.Add(buttonsPanel, 1, 0);

            // === НИЖНЯЯ ПАНЕЛЬ: СТАТУС ===
            var statusPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.DarkGray };
            var statusLabel = new Label
            {
                Text = "Панель отладки готова | Логирование активно",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.White,
                Padding = new Padding(10, 0, 0, 0)
            };
            statusPanel.Controls.Add(statusLabel);
            mainLayout.Controls.Add(statusPanel, 0, 1);
            mainLayout.SetColumnSpan(statusPanel, 2);

            this.Controls.Add(mainLayout);
        }

        private Button CreateButton(string text, Color backColor, Color foreColor)
        {
            return new Button
            {
                Text = text,
                Width = 200,
                Height = 40,
                BackColor = backColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(5)
            };
        }

        private void SetupEventHandlers()
        {
            LoggingSystem.LogAdded += OnLogAdded;
        }

        private void OnLogAdded(object sender, LogEntry entry)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => OnLogAdded(sender, entry)));
                return;
            }

            if (_isPaused) return;

            lock (_logsLock)
            {
                _allLogs.Add(entry);
            }

            // Фильтрация по поиску
            if (!string.IsNullOrEmpty(_lastSearchTerm))
            {
                if (!entry.Message.Contains(_lastSearchTerm, StringComparison.OrdinalIgnoreCase) &&
                    !entry.ClassName.Contains(_lastSearchTerm, StringComparison.OrdinalIgnoreCase) &&
                    !entry.MethodName.Contains(_lastSearchTerm, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            AppendLogEntry(entry);
        }

        private void AppendLogEntry(LogEntry entry)
        {
            if (!_colorMode)
            {
                _logTextBox.AppendText(entry.ToString() + Environment.NewLine);
                return;
            }

            // Цветовая подсветка
            Color color = Color.LightGreen; // DEBUG/INFO
            if (entry.Level == LogLevel.WARNING) color = Color.Yellow;
            else if (entry.Level == LogLevel.ERROR) color = Color.Orange;
            else if (entry.Level == LogLevel.CRITICAL) color = Color.Red;

            _logTextBox.SelectionStart = _logTextBox.TextLength;
            _logTextBox.SelectionLength = 0;
            _logTextBox.SelectionColor = color;
            _logTextBox.AppendText(entry.ToString() + Environment.NewLine);
            _logTextBox.SelectionColor = _logTextBox.ForeColor;

            // Автоскролл
            _logTextBox.ScrollToCaret();
        }

        // === ОБРАБОТЧИКИ КНОПОК ===

        private void BtnPanic_Click(object sender, EventArgs e)
        {
            // Кастомный диалог выбора режима отправки
            using (var dialog = new Form())
            {
                dialog.Text = "🔴 РЕЖИМ ПАНИКИ";
                dialog.Size = new Size(600, 400);
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ShowInTaskbar = false;
                dialog.Padding = new Padding(20);

                var label = new Label
                {
                    Text = "🔴 КНОПКА ПАНИКИ НАЖАТА!\n\nВыберите режим отправки:",
                    Location = new Point(30, 30),
                    Size = new Size(540, 60),
                    Font = new Font("Segoe UI", 12, FontStyle.Bold),
                    AutoSize = false
                };

                var btn1 = new Button
                {
                    Text = "1 - Комплексная отправка\n(все файлы из папки logs)",
                    Location = new Point(30, 110),
                    Size = new Size(540, 80),
                    DialogResult = DialogResult.Yes,
                    Font = new Font("Segoe UI", 11, FontStyle.Bold),
                    BackColor = Color.FromArgb(46, 139, 87),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                var btn2 = new Button
                {
                    Text = "2 - Текущая сессия\n(только логи текущей сессии)",
                    Location = new Point(30, 210),
                    Size = new Size(540, 80),
                    DialogResult = DialogResult.No,
                    Font = new Font("Segoe UI", 11, FontStyle.Bold),
                    BackColor = Color.FromArgb(52, 152, 219),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                var btnCancel = new Button
                {
                    Text = "Отменить",
                    Location = new Point(30, 310),
                    Size = new Size(540, 45),
                    DialogResult = DialogResult.Cancel,
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    BackColor = Color.Gray,
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };

                dialog.Controls.Add(label);
                dialog.Controls.Add(btn1);
                dialog.Controls.Add(btn2);
                dialog.Controls.Add(btnCancel);
                dialog.AcceptButton = btn1;
                dialog.CancelButton = btnCancel;

                var modeResult = dialog.ShowDialog(this);

                if (modeResult == DialogResult.Cancel)
                {
                    return; // Пользователь отменил
                }

                // Yes = 1 (Комплексная), No = 2 (Текущая сессия)
                bool isFullMode = (modeResult == DialogResult.Yes);

                try
                {
                var logsDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "ZontSpecExtractor_Logs");
                Directory.CreateDirectory(logsDirectory);

                // 1. Сохраняем дамп памяти
                var dumpPath = Path.Combine(
                    logsDirectory,
                    $"memory_dump_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.txt");

                var dump = GenerateMemoryDump();
                File.WriteAllText(dumpPath, dump);

                // 2. Сохраняем текущие логи в файл
                var logFilePath = Path.Combine(
                    logsDirectory,
                    $"log_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
                var currentLogContent = LoggingSystem.GetLogBuffer();
                if (!string.IsNullOrEmpty(currentLogContent))
                {
                    File.WriteAllText(logFilePath, currentLogContent);
                }
                else
                {
                    // Если буфер пуст, сохраняем через SaveToFile
                    LoggingSystem.SaveToFile();
                    // Получаем последний созданный файл лога
                    var logFiles = Directory.GetFiles(logsDirectory, "log_*.log");
                    if (logFiles.Length > 0)
                    {
                        logFilePath = logFiles.OrderByDescending(f => File.GetCreationTime(f)).First();
                    }
                }

                // 3. Собираем файлы для отправки в зависимости от режима
                var filesToSend = new List<string> { dumpPath };
                
                if (isFullMode)
                {
                    // КОМПЛЕКСНАЯ ОТПРАВКА: добавляем все файлы из папки
                    if (File.Exists(logFilePath))
                    {
                        filesToSend.Add(logFilePath);
                    }

                    // Добавляем все остальные файлы логов из папки (не старше 1 часа)
                    var allLogFiles = Directory.GetFiles(logsDirectory, "*.*")
                        .Where(f => 
                        {
                            var ext = Path.GetExtension(f).ToLower();
                            return ext == ".log" || ext == ".txt";
                        })
                        .Where(f => 
                        {
                            var fileTime = File.GetLastWriteTime(f);
                            return (DateTime.Now - fileTime).TotalHours <= 1; // Файлы не старше 1 часа
                        })
                        .Where(f => !filesToSend.Contains(f))
                        .ToList();
                    
                    filesToSend.AddRange(allLogFiles);
                }
                else
                {
                    // ТЕКУЩАЯ СЕССИЯ: только дамп и текущий лог
                    if (File.Exists(logFilePath))
                    {
                        filesToSend.Add(logFilePath);
                    }
                }

                // 4. Отправляем в Telegram
                var telegramLogger = GetTelegramLogger();
                if (telegramLogger != null)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            // Формируем информационное сообщение
                            var panicMessage = new StringBuilder();
                            panicMessage.AppendLine("🔴 КНОПКА ПАНИКИ НАЖАТА!");
                            panicMessage.AppendLine($"Время: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                            panicMessage.AppendLine($"ID процесса: {Process.GetCurrentProcess().Id}");
                            panicMessage.AppendLine();
                            panicMessage.AppendLine($"Режим: {(isFullMode ? "Комплексная отправка" : "Текущая сессия")}");
                            panicMessage.AppendLine();
                            panicMessage.AppendLine("=== КРАТКАЯ СВОДКА ===");
                            panicMessage.AppendLine(dump);
                            panicMessage.AppendLine();
                            panicMessage.AppendLine($"Всего файлов для отправки: {filesToSend.Count}");

                            // Отправляем сообщение и файлы
                            await telegramLogger.SendFilesAsync(filesToSend, panicMessage.ToString());

                            this.Invoke(new Action(() =>
                            {
                                MessageBox.Show(
                                    $"Паника обработана!\n\n" +
                                    $"Режим: {(isFullMode ? "Комплексная отправка" : "Текущая сессия")}\n" +
                                    $"Дамп сохранен: {Path.GetFileName(dumpPath)}\n" +
                                    $"Отправлено файлов в Telegram: {filesToSend.Count}",
                                    "Паника завершена", 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Information);
                            }));
                        }
                        catch (Exception ex)
                        {
                            this.Invoke(new Action(() =>
                            {
                                MessageBox.Show(
                                    $"Ошибка при отправке в Telegram: {ex.Message}\n\n" +
                                    $"Дамп сохранен в:\n{dumpPath}",
                                    "Ошибка отправки", 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Warning);
                            }));
                        }
                    });
                }
                else
                {
                    MessageBox.Show(
                        $"Дамп паники сохранен в:\n{dumpPath}\n\n" +
                        $"Режим: {(isFullMode ? "Комплексная отправка" : "Текущая сессия")}\n" +
                        $"Telegram не настроен. Файлы не отправлены.",
                        "Паника завершена", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Information);
                }

                    // 5. Принудительная остановка (опционально)
                    // Application.Exit(); // Раскомментируйте если нужно
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка обработчика паники: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string GenerateMemoryDump()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== ДАМП ПАМЯТИ ===");
            sb.AppendLine($"Время: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"ID процесса: {Process.GetCurrentProcess().Id}");
            
            var metrics = LoggingSystem.GetSystemMetrics();
            sb.AppendLine($"Память: {metrics.MemoryBytes / 1024 / 1024} МБ");
            sb.AppendLine($"CPU: {metrics.CpuPercent:F2}%");
            
            sb.AppendLine($"Потоки: {Process.GetCurrentProcess().Threads.Count}");
            sb.AppendLine($"Сборки мусора: Gen0={GC.CollectionCount(0)}, Gen1={GC.CollectionCount(1)}, Gen2={GC.CollectionCount(2)}");
            
            sb.AppendLine("\n=== СВОДКА ЛОГОВ ===");
            lock (_logsLock)
            {
                sb.AppendLine($"Всего записей: {_allLogs.Count}");
                sb.AppendLine($"Ошибки: {_allLogs.Count(l => l.Level >= LogLevel.ERROR)}");
                sb.AppendLine($"Предупреждения: {_allLogs.Count(l => l.Level == LogLevel.WARNING)}");
            }
            
            return sb.ToString();
        }

        private void BtnSendTG_Click(object sender, EventArgs e)
        {
            var telegramLogger = GetTelegramLogger();
            if (telegramLogger == null)
            {
                MessageBox.Show("Логгер Telegram не настроен. Установите BOT_TOKEN и CHAT_ID в настройках.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var logContent = LoggingSystem.GetLogBuffer();
            if (string.IsNullOrEmpty(logContent))
            {
                MessageBox.Show("Нет логов для отправки.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _ = Task.Run(async () =>
            {
                await telegramLogger.SendFullLogAsync(logContent);
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show("Лог отправлен в Telegram!", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            });
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Файлы логов (*.log)|*.log|Все файлы (*.*)|*.*";
                sfd.FileName = $"log_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log";
                sfd.Title = "Сохранить лог";
                
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        File.WriteAllText(sfd.FileName, LoggingSystem.GetLogBuffer());
                        MessageBox.Show($"Лог сохранен в:\n{sfd.FileName}", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось сохранить: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnOpenLogsFolder_Click(object sender, EventArgs e)
        {
            try
            {
                var logsDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "ZontSpecExtractor_Logs");

                // Создаем папку, если её нет
                if (!Directory.Exists(logsDirectory))
                {
                    Directory.CreateDirectory(logsDirectory);
                }

                // Открываем папку в проводнике Windows
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = logsDirectory,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть папку с логами: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AutoSaveCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (_autoSaveCheckBox.Checked)
            {
                int interval = (int)_autoSaveInterval.Value;
                LoggingSystem.EnableAutoSave(interval);
            }
            else
            {
                LoggingSystem.DisableAutoSave();
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            _logTextBox.Clear();
            lock (_logsLock)
            {
                _allLogs.Clear();
            }
            LoggingSystem.ClearLogBuffer();
        }

        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            _lastSearchTerm = _searchTextBox.Text;
            
            // Перерисовываем логи с фильтром
            _logTextBox.Clear();
            lock (_logsLock)
            {
                var filtered = string.IsNullOrEmpty(_lastSearchTerm)
                    ? _allLogs
                    : _allLogs.Where(log =>
                        log.Message.Contains(_lastSearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        log.ClassName.Contains(_lastSearchTerm, StringComparison.OrdinalIgnoreCase) ||
                        log.MethodName.Contains(_lastSearchTerm, StringComparison.OrdinalIgnoreCase)).ToList();

                foreach (var entry in filtered)
                {
                    AppendLogEntry(entry);
                }
            }
        }

        private void BtnPause_Click(object sender, EventArgs e)
        {
            _isPaused = !_isPaused;
            if (_isPaused)
            {
                LoggingSystem.Pause();
                ((Button)sender).Text = "▶ Возобновить логирование";
                ((Button)sender).BackColor = Color.Green;
            }
            else
            {
                LoggingSystem.Resume();
                ((Button)sender).Text = "⏸ Приостановить логирование";
                ((Button)sender).BackColor = Color.Yellow;
            }
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(_logTextBox.Text);
                MessageBox.Show("Лог скопирован в буфер обмена!", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось скопировать: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnEmail_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Функция Email отчета является заглушкой.\n" +
                "Для реализации:\n" +
                "1. Настроить SMTP параметры\n" +
                "2. Добавить шаблон письма\n" +
                "3. Реализовать логику вложений",
                "Email отчет",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void BtnColor_Click(object sender, EventArgs e)
        {
            _colorMode = !_colorMode;
            if (_colorMode)
            {
                _logTextBox.BackColor = Color.Black;
                _logTextBox.ForeColor = Color.LightGreen;
            }
            else
            {
                _logTextBox.BackColor = Color.White;
                _logTextBox.ForeColor = Color.Black;
            }
            
            // Перерисовываем все логи
            _logTextBox.Clear();
            lock (_logsLock)
            {
                foreach (var entry in _allLogs)
                {
                    AppendLogEntry(entry);
                }
            }
        }

        private void BtnTestError_Click(object sender, EventArgs e)
        {
            try
            {
                LoggingSystem.Log(LogLevel.INFO, "DebugPanel", "BtnTestError_Click", "Injecting test error...");
                throw new Exception("🐛 TEST ERROR: This is an artificially injected error for testing the logging system!");
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("DebugPanel", "BtnTestError_Click", ex);
            }
        }

        private void StartMetricsTimer()
        {
            _metricsTimer = new System.Windows.Forms.Timer();
            _metricsTimer.Interval = 2000; // 2 секунды
            _metricsTimer.Tick += (sender, e) => UpdateMetrics();
            _metricsTimer.Start();
        }

        private void UpdateMetrics()
        {
            try
            {
                var metrics = LoggingSystem.GetSystemMetrics();
                _memoryLabel.Text = $"Память: {metrics.MemoryBytes / 1024 / 1024} МБ";
                _cpuLabel.Text = $"CPU: {metrics.CpuPercent:F2}%";
            }
            catch { }
        }

        private static TelegramLogger _telegramLoggerInstance = null;
        private static readonly object _telegramLock = new object();

        private TelegramLogger GetTelegramLogger()
        {
            if (_telegramLoggerInstance != null) return _telegramLoggerInstance;
            
            lock (_telegramLock)
            {
                if (_telegramLoggerInstance != null) return _telegramLoggerInstance;
                
                // Получаем из настроек
                const string token = "8274395823:AAFyn_uRp6jhNnbbSKoT74EuSWFiIedAVVw";
                const string chatId = "1038655823";
                
                _telegramLoggerInstance = new TelegramLogger(token, chatId);
                return _telegramLoggerInstance;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _metricsTimer?.Dispose();
            LoggingSystem.LogAdded -= OnLogAdded;
            base.OnFormClosing(e);
        }
    }
}


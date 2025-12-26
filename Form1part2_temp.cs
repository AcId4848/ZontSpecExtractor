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

    public partial class Form1
    {
        private void CreateTableClick(object? sender, EventArgs e)
        {
            _lastActivityTime = DateTime.Now; // ��������� ����������
            if (!_rawHits.Any())
            {
                MessageBox.Show("������� ���������� ��������� � ���������������� Excel-����!");
                return;
            }

            // =================================================================================
            // 1. Конфигурация и настройки (основные классы)���������� ������ (������� �������� � ������� Excel)
            // =================================================================================

            // �� ���������� ������� �� ����� ����� � ������ ������.
            // ��� �����������, ��� ������ ����� ������ ��������� ������.
            var rowsData = _rawHits
                .Where(h => h.ConditionMet) // ����� ������ ��, ��� ����������� ������� (��� ���� "1")
                .GroupBy(h => new { h.SheetName, h.RowIndex }) // ����������� �� ������
                .Select(g => new
                {
                    // ����������: ������� �� ������, ����� �� ������ ������ � Excel
                    SortSheet = g.Key.SheetName,
                    SortRow = g.Key.RowIndex,

                    // ���� ������ ��� ������ ������� ������ � ���� ������ (g)

                    // �������� ������� (TargetColumn = "Position" ��� �����)
                    Position = g.FirstOrDefault(x => string.IsNullOrEmpty(x.TargetColumn) || x.TargetColumn == "Position")?.FullItemName,

                    // ���������� (����� �� ������� ��� 1)
                    Quantity = g.Where(x => string.IsNullOrEmpty(x.TargetColumn) || x.TargetColumn == "Position").Sum(x => x.Quantity),

                    // ��������� �������. ���� � ���� ������ ��� ������ ��� Col4, ����� null (�������)
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
                // ����������: ������ ��� � Excel (�� ������� �����)
                .OrderBy(x => x.SortSheet).ThenBy(x => x.SortRow)
                .ToList();

            if (!rowsData.Any())
            {
                MessageBox.Show("�� ������� ������ ��� ������������ �������.");
                return;
            }

            // =================================================================================
            // 2. Основной класс формы (GeneralSettingsForm)�������� � ��������� �������
            // =================================================================================

            // ���������� ��������� ����� ����� (����������� ������ CS0029)
            string visioSheetName = "";
            try
            {
                var source = AppSettings.SearchConfig.VisioSourceSheetName;
                if (source is IEnumerable<string> list) visioSheetName = list.FirstOrDefault() ?? "";
                else visioSheetName = source?.ToString() ?? "";
            }
            catch { visioSheetName = "������ �����"; }

            var tableForm = new Form
            {
                Text = $"������������ (���� Visio: {(string.IsNullOrEmpty(visioSheetName) ? "�� ������" : visioSheetName)})",
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

            // --- ���������� ������� ---
            dgv.Columns.Add("Priority", "����� ����������");
            dgv.Columns.Add("Position", "������� (��������)"); dgv.Columns["Position"].Width = 200;
            dgv.Columns.Add("Quantity", "����������");

            dgv.Columns.Add("Col4", "����/�����/����");
            dgv.Columns.Add("Col5", "��. �����/����");
            dgv.Columns.Add("Col6", "���������/����2/������");
            dgv.Columns.Add("Col7", "��. ������/���� ����");
            dgv.Columns.Add("Col8", "��. ������/���� ����");
            dgv.Columns.Add("Col9", "��. 1 ���� ������");
            dgv.Columns.Add("Col10", "��. 2 ���� ������");
            dgv.Columns.Add("Col11", "�����");
            dgv.Columns.Add("Col12", "������� 1");
            dgv.Columns.Add("Col13", "������2 / ������� +5�");
            dgv.Columns.Add("Col14", "��. 1 ������ (-12�)");
            dgv.Columns.Add("Col15", "��. 2 ������ T1");
            dgv.Columns.Add("Col16", "��. 3 ������ T2");
            dgv.Columns.Add("Col17", "��. 4 ������ ������");
            dgv.Columns.Add("Col18", "��. 5 ������ (+5�)");

            // =================================================================================
            // 3. ���������� ������� (������ �� �������)
            // =================================================================================

            for (int i = 0; i < rowsData.Count; i++)
            {
                var dataItem = rowsData[i];

                int rowIndex = dgv.Rows.Add();
                var row = dgv.Rows[rowIndex];

                // ����� �/� (������ ���������� ����� � �������� �������)
                row.Cells["Priority"].Value = i + 1;

                // �������
                row.Cells["Position"].Value = dataItem.Position;

                // ���������� (���� 0, ����� �� ������ ��� ������ 0 - �� �������)
                // ���� ������� ������, �� ���� �����, ���������� ����� ���� 0, ���� �� ��������� �����
                row.Cells["Quantity"].Value = dataItem.Quantity == 0 ? 1 : dataItem.Quantity;

                // ��������� ��������� �������. 
                // ���� � dataItem.Col4 ������ ��� (null), � ������ ��������� null (�������).
                // ��������� �������� �� ����������� ����.
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

        /// <summary>
        /// ���������� ������ �������� ������� ������� � Excel
        /// </summary>
        private void BtnExportExcel_Click(object? sender, EventArgs e)
        {
            _lastActivityTime = DateTime.Now; // ��������� ����������
            try
            {
                if (dataGridView == null || dataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("��� ������ ��� ��������. ������� ��������� � ����������� ����.", 
                        "��� ������", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ExportDataGridViewToExcel(dataGridView);
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("Form1", "BtnExportExcel_Click", ex);
                MessageBox.Show($"������ ��� �������� � Excel: {ex.Message}", 
                    "������ ��������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ������������� ����� �������� DataGridView � Excel
        /// </summary>
        private void ExportDataGridViewToExcel(DataGridView dgv)
        {
            try
            {
                string outputPath;
                string defaultFileName = "�������.xlsx";
                
                // ���������� ������������ ��� �����
                if (!string.IsNullOrEmpty(AppSettings.LastLoadedFilePath))
                {
                    var sourceFileInfo = new FileInfo(AppSettings.LastLoadedFilePath);
                    var sourceFileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceFileInfo.Name);
                    defaultFileName = $"{sourceFileNameWithoutExt} [�������].xlsx";
                }

                // ������ ���������� ������ ����������
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveDialog.FileName = defaultFileName;
                    saveDialog.Title = "��������� ������� � Excel";
                    
                    // ������������� ��������� ����������
                    if (!string.IsNullOrEmpty(AppSettings.LastLoadedFilePath))
                    {
                        var sourceFileInfo = new FileInfo(AppSettings.LastLoadedFilePath);
                        saveDialog.InitialDirectory = sourceFileInfo.DirectoryName ?? Environment.CurrentDirectory;
                    }
                    
                    if (saveDialog.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }
                    outputPath = saveDialog.FileName;
                }

                // ������� Excel ����
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("�������");

                    // ���������� ���������
                    for (int col = 0; col < dgv.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = dgv.Columns[col].HeaderText;
                        worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                        worksheet.Cells[1, col + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[1, col + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // ���������� ������
                    for (int row = 0; row < dgv.Rows.Count; row++)
                    {
                        if (dgv.Rows[row].IsNewRow) continue;
                        
                        for (int col = 0; col < dgv.Columns.Count; col++)
                        {
                            var cellValue = dgv.Rows[row].Cells[col].Value;
                            if (cellValue != null)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = cellValue.ToString();
                            }
                        }
                    }

                    // ���������� ������ �������
                    if (worksheet.Dimension != null)
                    {
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    }

                    // ��������� ����
                    var fileInfo = new FileInfo(outputPath);
                    package.SaveAs(fileInfo);

                    LoggingSystem.Log(LogLevel.INFO, "Form1", "ExportDataGridViewToExcel", 
                        $"Table exported to: {outputPath}");
                    
                    UpdateStatus($"������� ��������������: {Path.GetFileName(outputPath)}");
                    MessageBox.Show($"������� ������� �������������� �:\n{outputPath}", 
                        "������� ��������", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("Form1", "ExportDataGridViewToExcel", ex);
                throw;
            }
        }

        /// <summary>
        /// ������������ ������� � Excel ���� (��� ������� �������� �������)
        /// </summary>
        private void ExportTableToExcel(DataGridView dgv, object rowsData)
        {
            try
            {
                if (string.IsNullOrEmpty(AppSettings.LastLoadedFilePath))
                {
                    LoggingSystem.Log(LogLevel.WARNING, "Form1", "ExportTableToExcel", 
                        "LastLoadedFilePath is empty, cannot determine output file name");
                    return;
                }

                // ��������� ��� �����: �������� ��� + [�������].xlsx
                var sourceFileInfo = new FileInfo(AppSettings.LastLoadedFilePath);
                var sourceFileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceFileInfo.Name);
                var outputDirectory = sourceFileInfo.DirectoryName ?? Environment.CurrentDirectory;
                var outputFileName = $"{sourceFileNameWithoutExt} [�������].xlsx";
                var outputPath = Path.Combine(outputDirectory, outputFileName);

                // ���� ���� ��� ����������, ��������� �����
                int counter = 1;
                while (File.Exists(outputPath))
                {
                    outputFileName = $"{sourceFileNameWithoutExt} [�������] ({counter}).xlsx";
                    outputPath = Path.Combine(outputDirectory, outputFileName);
                    counter++;
                }

                // ������� Excel ����
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("�������");

                    // ���������� ���������
                    for (int col = 0; col < dgv.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = dgv.Columns[col].HeaderText;
                        worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                        worksheet.Cells[1, col + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[1, col + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // ���������� ������
                    for (int row = 0; row < dgv.Rows.Count; row++)
                    {
                        for (int col = 0; col < dgv.Columns.Count; col++)
                        {
                            var cellValue = dgv.Rows[row].Cells[col].Value;
                            if (cellValue != null)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = cellValue.ToString();
                            }
                        }
                    }

                    // ���������� ������ �������
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // ��������� ����
                    var fileInfo = new FileInfo(outputPath);
                    package.SaveAs(fileInfo);

                    LoggingSystem.Log(LogLevel.INFO, "Form1", "ExportTableToExcel", 
                        $"Table exported to: {outputPath}");
                    
                    UpdateStatus($"������� ��������������: {outputFileName}");
                }
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("Form1", "ExportTableToExcel", ex);
                throw;
            }
        }

        private async void LoadFiles(string[] filePaths)
        {
            MethodLogger.LogEntry(string.Join(", ", filePaths ?? new string[0]));
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                LoggingSystem.Log(LogLevel.INFO, "Form1", "LoadFiles", 
                    $"Starting file load process for {filePaths?.Length ?? 0} file(s)");
                
                // ������� ���������� ����� �������
                dataGridView.Rows.Clear();
                data.Clear();
                _rawHits.Clear();

                if (filePaths.Length > 0)
                {
                    var fileName = System.IO.Path.GetFileName(filePaths[0]);
                    lblFileInfo.Text = $"📄 {fileName}";
                    // Обновляем заголовок окна с полным путем к файлу
                    UpdateWindowTitle(filePaths[0]);
                }
                else
                {
                    lblFileInfo.Text = "Файл не выбран";
                    LoggingSystem.Log(LogLevel.WARNING, "Form1", "LoadFiles", "No file paths provided");
                    return;
                }

                UpdateStatus("? ���� ������������ Excel-�����...");

                try
                {
                    await Task.Run(() =>
                    {
                        foreach (var path in filePaths)
                        {
                            LoggingSystem.Log(LogLevel.INFO, "Form1", "LoadFiles", 
                                $"Scanning file: {path}");
                            
                            // ��������� ������ �� �����, ������� ������� � ������ ������� 
                            // (ScanSpecificSheet ����� �� �� AppSettings.SearchConfig.TargetSheetNames)
                            var hits = ScanSpecificSheet(path);
                            _rawHits.AddRange(hits);
                            
                            LoggingSystem.Log(LogLevel.INFO, "Form1", "LoadFiles", 
                                $"Found {hits.Count} hits in file {System.IO.Path.GetFileName(path)}");
                        }
                    });

                    // 1. Конфигурация и настройки (основные классы)���������� ����� ������ (_rawHits) � ������ ��� ��������� ���������
                    data = GroupRawHitsForVisio(_rawHits);

                    // 2. Основной класс формы (GeneralSettingsForm)��������� ������� �� ������� �����
                    UpdateDataGridView();
                    ShowResultMessage(_rawHits.Count);

                    stopwatch.Stop();
                    LoggingSystem.Log(LogLevel.INFO, "Form1", "LoadFiles", 
                        $"File load completed in {stopwatch.ElapsedMilliseconds}ms. Total hits: {_rawHits.Count}");

                    // ����������: _visioSourceSheetName ������ ������ �������� � ������ 
                    // � ����� ����������� ��� ������� ������ "������� �������".
                }
                catch (Exception ex)
                {
                    stopwatch.Stop();
                    LoggingSystem.LogException("Form1", "LoadFiles", ex, 
                        new Dictionary<string, object> { { "FilePaths", string.Join(", ", filePaths) } });
                    MessageBox.Show($"������ ��� ������ �����: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus("������ ��������.");
                }
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("Form1", "LoadFiles", ex);
                throw;
            }
            finally
            {
                MethodLogger.LogExit();
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
                MessageBox.Show("�� ������� ������ ��� ������������ �������.");
                return;
            }

            // ������� �������� ����������
            double GetPriority(string? val)
            {
                if (string.IsNullOrEmpty(val)) return 999999.0;
                if (double.TryParse(val.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double res))
                    return res;
                return 999999.0;
            }

            // ���� ������ �� ������ ������� � ���� ������
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
            // 2. Основной класс формы (GeneralSettingsForm)��������� ������� (������ ������������ �������)
            // =================================================================================
            mainDgv = this.dataGridView;
            if (mainDgv != null)
            {
                mainDgv.AutoGenerateColumns = false;
                mainDgv.Columns.Clear();

                // --- �������� ��� ������ ��� ����������� ---
                // 1. Конфигурация и настройки (основные классы)��������� ����������� (EnableAlwaysIncludeHeaderText �������� � ���������, ������ ��� Excel)
                mainDgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

                // 2. Основной класс формы (GeneralSettingsForm)��������� �������� ��������� ����� ��� �����
                mainDgv.MultiSelect = true;

                // 3. ����� ���������: RowHeaderSelect ��������� �������� � ������ ������� (����� �����), � ��������� ������
                mainDgv.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;        // ������� ��, ��� ���� ������

                // ������� 5 ������ �������
                mainDgv.Columns.Add("BlockNum", "�");
                mainDgv.Columns.Add("Terminals", "������");
                mainDgv.Columns.Add("BlockName", "������������ �����");
                mainDgv.Columns.Add("SourcePos", "������� �� ��");
                mainDgv.Columns.Add("PriorityVal", "���������");

                // ��������� ������ (��� �������)
                mainDgv.Columns["BlockNum"].Width = 40;
                mainDgv.Columns["Terminals"].Width = 80;
                mainDgv.Columns["BlockName"].Width = 250;
                mainDgv.Columns["SourcePos"].Width = 60;
                mainDgv.Columns["PriorityVal"].Width = 60;

                mainDgv.Rows.Clear();

                // =================================================================================
                // 3. �������� ���� ����������
                // =================================================================================

                // ����� ������� ���������� ������./������. ��� ������� ������ ������ ������
                int lastSwitchIndex = finalQueue.FindLastIndex(x =>
                    (x.Name ?? "").Contains("������.", StringComparison.OrdinalIgnoreCase) ||
                    (x.Name ?? "").Contains("������.", StringComparison.OrdinalIgnoreCase)
                );

                // ���������, ���� �� ������ "������." � ������ (��� ������ �������)
                bool hasPereklForFooter = finalQueue.Any(x => (x.Name ?? "").Contains("������.", StringComparison.OrdinalIgnoreCase));
                bool hasForFooter = finalQueue.Any(x => (x.Name ?? "").Contains("������.", StringComparison.OrdinalIgnoreCase));

                int currentTerminalCounter = 3; // �������� � 3-� ������
                int sensorsInGroup = 0;         // ������� ������� ������ �������� (0, 1, 2...)

                // �������: ��� ������?
                bool IsSensorRow(string name)
                {
                    if (string.IsNullOrEmpty(name)) return false;
                    return name.Contains("�", StringComparison.Ordinal) && !(name.Contains("�/", StringComparison.Ordinal)) ||
                           name.Contains("��", StringComparison.Ordinal);
                }

                string GetSmartBoilerName(string currentBlockName, string sourcePosName)
                {
                    if (string.IsNullOrWhiteSpace(sourcePosName) || string.IsNullOrWhiteSpace(currentBlockName))
                        return currentBlockName;

                    try
                    {
                        // ���� �������: ����� "�����", �������, ����� "1", �������, (��� ������), (����)
                        // RegexOptions.IgnoreCase - ����� �� �������� �� ��������
                        var regex = new System.Text.RegularExpressions.Regex(
                            @"�����\s+1\s+(.*?)\(?����\)?",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        var match = regex.Match(sourcePosName);

                        if (match.Success)
                        {
                            // ������ 1 - ��� ��, ��� ������ ����� "����� 1" � "����"
                            string modelName = match.Groups[1].Value.Trim();

                            if (!string.IsNullOrEmpty(modelName))
                            {
                                // �������� ����� "1" ��� ��������� ����� (\b1\b) �� �������� ������
                                // ��� ������������� ������ ������ ������ ������ ����� (��������, 12� �� ����������)
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

                    // 1. Конфигурация и настройки (основные классы)�������� "�����" ���
                    // ���� � �������� ������� (SourcePos) ���� ������ �����, ��� ����������� ������ "1"
                    string originalName = item.Name ?? "";
                    string name = GetSmartBoilerName(originalName, item.SourcePos ?? "");

                    bool isSensor = IsSensorRow(name);

                    // --- 3.1 ������� ������ (����� ���������) ---
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

                    // --- 3.2 ��������� �������� ������ ---
                    int r = mainDgv.Rows.Add();
                    var row = mainDgv.Rows[r];

                    // ������ ����� (���������� ��� ����� ��� name)
                    string terminalValue = "";
                    if (name.Contains("������� 220�", StringComparison.OrdinalIgnoreCase))
                    {
                        terminalValue = "1";
                    }
                    else if (name.Contains("���������", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("������.", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("������.", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("�. ��������", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("����", StringComparison.OrdinalIgnoreCase) ||
                             name.Contains("������", StringComparison.OrdinalIgnoreCase))
                    {
                        terminalValue = $"{currentTerminalCounter}, {currentTerminalCounter + 1}";
                        currentTerminalCounter += 2;
                    }
                    else
                    {
                        terminalValue = currentTerminalCounter.ToString();
                        currentTerminalCounter += 1;
                    }

                    // ����������
                    row.Cells["BlockNum"].Value = r + 1;
                    row.Cells["Terminals"].Value = terminalValue;

                    // �����: ���������� � ������� ����� ��� � Buderus
                    row.Cells["BlockName"].Value = name;

                    row.Cells["SourcePos"].Value = item.SourcePos;
                    row.Cells["PriorityVal"].Value = item.PriorityRaw;

                    // --- 3.3 ���������� ������������ �������� ---
                    if (isSensor)
                    {
                        sensorsInGroup++;
                        int sensorsAhead = 0;
                        // ��� �������� ��������� ����� ��� �� ����� ������ �� �����, 
                        // ���������� ��������� �������� IsSensorRow, ��� ��� "�" ��� "��" �� �������� ��� ������ �����
                        for (int k = i + 1; k < finalQueue.Count; k++)
                        {
                            string nextName = finalQueue[k].Name ?? ""; // ����� ����� ��� ��� �������� ����
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
                // 4. "������" ������� (���� ���� "������.")
                // =================================================================================
                if (hasPereklForFooter || hasForFooter)
                {
                    // 1. Конфигурация и настройки (основные классы)����� -12� (GND)
                    AddRowManual(mainDgv, ref currentTerminalCounter, "����� -12� (GND)", 1);

                    // 2. Основной класс формы (GeneralSettingsForm)������� +12�
                    AddRowManual(mainDgv, ref currentTerminalCounter, "������� +12�", 1);

                    // 3. ���� ����� RS485 (�������� 2 ������!)
                    AddRowManual(mainDgv, ref currentTerminalCounter, "���� ����� RS485", 2);
                }

                mainDgv.Refresh();
            }
            else
            {
                MessageBox.Show("�� ������� ����� ������� �������.");
            }
        }

        // =================================================================================
        // ��������������� ������
        // =================================================================================

        // 1. Конфигурация и настройки (основные классы)����� ��� ������� �������� "���������������" ������
        private void AddMinusRow(DataGridView dgv, ref int terminalCounter)
        {
            int r = dgv.Rows.Add();
            var row = dgv.Rows[r];

            // �������� 1 ������
            string term = terminalCounter.ToString();
            terminalCounter += 1;

            row.Cells["BlockNum"].Value = r + 1;
            row.Cells["Terminals"].Value = term;
            row.Cells["BlockName"].Value = "����� -12� (GND)";
            row.Cells["SourcePos"].Value = "";
            row.Cells["PriorityVal"].Value = "";
        }

        // 2. Основной класс формы (GeneralSettingsForm)������������� ����� ��� ���������� ������ � ����� (��� RS485 � �������)
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
            else // ���� 2 ������
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

        // ��������������� ����� ��� ���������� ������ � ������� (����� �� ����������� ���)

        private List<Dictionary<string, string>> GroupRawHitsForVisio(List<RawExcelHit> rawHits)
        {
            // ��� ������ - ��� ��, ��� ���� ������� �� ScanSpecificSheet. 
            // ��� ���������� �� �������� ����� (SearchTerm), ��� ����������� ��� Visio �����.

            return rawHits
                .GroupBy(h => h.SearchTerm) // ���������� �� �������� �����
                .Select(g =>
                {
                    // ���������, ���� �� � ����� ������� ����������� (������� �� ���� ������ ���������� ������ ������)
                    bool isLimited = g.First().IsLimited;

                    int totalQty;
                    if (isLimited)
                    {
                        // ���� ����� ������� "����������", �� ���������� �� ���������� ��������� �����, ����� = 1
                        totalQty = 1;
                    }
                    else
                    {
                        // ����� ��������� ��, ��� �����
                        totalQty = g.Sum(x => x.Quantity);
                    }

                    return new Dictionary<string, string>
                    {
                        ["����"] = g.First().SheetName,
                        ["������������"] = g.Key,
                        ["����������"] = totalQty.ToString()
                    };
                })
                .Where(x => x["����������"] != "0")
                .ToList();
        }

        //private void btnApplySettings_Click(object sender, EventArgs e)
        //{
        //    // 1. Конфигурация и настройки (основные классы)��������� �����
        //    AppSettings.SearchConfig.TargetSheetNames = _textBoxTargetSheets.Text
        //        .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
        //        .Select(s => s.Trim())
        //        .ToList();

        //    // <-- ����������� ����������� ���������� ���� -->
        //    // 2. Основной класс формы (GeneralSettingsForm)��������� ����� ��� ������
        //    AppSettings.SearchConfig.SearchWords = _textBoxSearchWords.Text
        //        .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
        //        .Select(s => s.Trim())
        //        .ToList();
        //    // ----------------------------------------------

        //    // 3. ��������� ��� ��������� �� ����
        //    AppSettings.Save();

        //    UpdateStatus("��������� ������� ��������� � ���������.");
        //}
        //    // -----------------------------------------------

        //    UpdateStatus("��������� ������� ��������� � ���������.");
        //}

        // --- ����������� �����: ����� � ������� (������� ������� �� ������) ---
        // --- ����� �����: ����� � ����� ����������� ����� (������ ���� �� ������) ---
        // --- ����������� �����: �����, ������� � ��������� �� ��������� ������ ---

        /// <summary>
        /// ������� ���������� ��������� ��������� � ������ ��� ����� ��������.
        /// </summary>
        private static int CountOccurrences(string source, string word)
        {
            if (string.IsNullOrEmpty(word) || string.IsNullOrEmpty(source))
            {
                return 0;
            }

            int count = 0;
            int index = -1;
            // ���������� IndexOf � StringComparison.OrdinalIgnoreCase ��� ������ ��� ����� ��������
            // � ���������� ���� ��� ���������.
            while ((index = source.IndexOf(word, index + 1, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                count++;
            }
            return count;
        }

        public class ExcelMatch
        {
            public string Sheet { get; set; } = ""; // <- ��������� ��� ���������� CS0649
            public string Value { get; set; } = ""; // <- ��������� ��� ���������� CS0649

        }

        // --- ����� �����: �����, ����� ����������� � ������� ������ ����� ---
        // --- ����� �����: �����, ����� ����������� � ������� ������ ����� ---
        // ��������� 1: ������ ������������ ���
        // =============================================================
        // ����� 1: �������� "����� �����" (�����������) �� ����. �����
        // =============================================================
        private void LoadPrioritiesFromVisioSheet(OfficeOpenXml.ExcelPackage package)
        {
            _visioPriorityMap.Clear();

            // 1. Конфигурация и настройки (основные классы)���� ���� � ������������ (�� ��������� ����� VISIO)
            var sheet = package.Workbook.Worksheets["1.�� �� ������ ZONT ��� VISIO"];
            if (sheet == null)
            {
                sheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.ToUpper().Contains("VISIO"));
            }

            if (sheet == null) return; // ����� ��� � ���������� �� ��������, �� �������

            int startRow = 2; // ���������� �����
            int endRow = sheet.Dimension?.End.Row ?? 100;

            for (int r = startRow; r <= endRow; r++)
            {
                // ������� 1 (A): ������������ (����. "����� ���������")
                // ������� 2 (B): ����� ������ (����. "3" ��� "������ 3")
                string name = sheet.Cells[r, 1].Text.Trim().ToLower();
                string valStr = sheet.Cells[r, 2].Text.Trim();

                if (string.IsNullOrEmpty(name)) continue;

                int priority = 9999;

                // ����������� ������ ����� �� ������ (����� "��. 3" ������������ � 3)
                string digits = new string(valStr.Where(char.IsDigit).ToArray());

                if (int.TryParse(digits, out int p))
                {
                    priority = p;
                }
                else
                {
                    // ���� ������ ���, ����� ����� ������ * 100 (����� ��������� ������� �������)
                    priority = r * 100;
                }

                if (!_visioPriorityMap.ContainsKey(name))
                {
                    _visioPriorityMap.Add(name, priority);
                }
            }
        }

        // =============================================================
        // ����� 2: �������� ������ (� ���������� ������� SchematicOrder)
        // =============================================================

        private bool IsRowExcluded(int currentRow, string exclusionRule)
        {
            if (string.IsNullOrWhiteSpace(exclusionRule)) return false;

            // ��������� ������ �� �������
            var parts = exclusionRule.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                string p = part.Trim();

                // ��������� ������ "229-" (�� ����� � �� �����)
                if (p.EndsWith("-"))
                {
                    string numberPart = p.TrimEnd('-');
                    if (int.TryParse(numberPart, out int startLimit))
                    {
                        if (currentRow >= startLimit) return true; // ���������, ���� ������ ������ ��� �����
                    }
                }
                // ��������� ������ "10" (���������� ������)
                else
                {
                    if (int.TryParse(p, out int specificRow))
                    {
                        if (currentRow == specificRow) return true; // ��������� ���������� ������
                    }
                }
            }
            return false;
        }

        private List<RawExcelHit> ScanSpecificSheet(string filePath)
        {
            MethodLogger.LogEntry(filePath);
            var rawHits = new List<RawExcelHit>();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            try
            {
                LoggingSystem.Log(LogLevel.DEBUG, "Form1", "ScanSpecificSheet", 
                    $"Starting scan of file: {filePath}");
                
                // 1. Конфигурация и настройки (основные классы)��������� �������� ���� (��� ��������� �������)
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    try 
                    { 
                        LoadPrioritiesFromVisioSheet(package);
                        LoggingSystem.Log(LogLevel.DEBUG, "Form1", "ScanSpecificSheet", "Priorities loaded from Visio sheet");
                    } 
                    catch (Exception ex)
                    {
                        LoggingSystem.Log(LogLevel.WARNING, "Form1", "ScanSpecificSheet", 
                            $"Failed to load priorities: {ex.Message}", ex);
                    }

                    // 2. Основной класс формы (GeneralSettingsForm)��������� �������������� ���� (������ ����� ��������)
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
                            // ����� ������� ���������, ���� ���� �����
                            System.Diagnostics.Debug.WriteLine($"������ �������� ���. �����: {ex.Message}");
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
                            // --- �������� ���� (��� �������) ---
                            var mainSheet = package.Workbook.Worksheets[sheetName];
                            if (mainSheet == null || mainSheet.Dimension == null) continue;

                            // --- �������������� ���� (��� ������) ---
                            ExcelWorksheet auxSheet = null;
                            if (auxPackage != null)
                            {
                                // ���� ���� ��� ����� �������� (Sheet1 == sheet1)
                                auxSheet = auxPackage.Workbook.Worksheets
                                    .FirstOrDefault(w => w.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                            }

                            int startRow = mainSheet.Dimension.Start.Row;
                            int endRow = mainSheet.Dimension.End.Row;

                            // === ���� �� ������� ===
                            for (int row = startRow; row <= endRow; row++)
                            {
                                foreach (var rule in AppSettings.SearchConfig.Rules)
                                {
                                    // ������� ����������� �����
                                    if (IsRowExcluded(row, rule.ExcludedRows)) continue;

                                    bool matchFound = false;
                                    string foundInMainTable = ""; // ��� ������� ��� ��������, ���� �����

                                    // =========================================================
                                    // 1. Конфигурация и настройки (основные классы)�������� ������� (������ � �������� �������)
                                    // =========================================================

                                    // ������� �: ����� �� �������� (��������, "1" � ������� L)
                                    if (rule.SearchByValue)
                                    {
                                        int condColIndex = ExcelColumnLetterToNumber(rule.ConditionColumn);
                                        if (condColIndex > 0)
                                        {
                                            // ������� � MAIN SHEET
                                            string cellValue = mainSheet.Cells[row, condColIndex].Text?.Trim();
                                            string targetValue = string.IsNullOrEmpty(rule.ConditionValue) ? "1" : rule.ConditionValue;

                                            if (string.Equals(cellValue, targetValue, StringComparison.OrdinalIgnoreCase))
                                            {
                                                matchFound = true;
                                                // ����������, ��� ����� � �������� (�� ������ ������), �� ���� �� ����������
                                                int nameCol = ExcelColumnLetterToNumber(rule.SearchColumn);
                                                if (nameCol > 0) foundInMainTable = mainSheet.Cells[row, nameCol].Text?.Trim();
                                            }
                                        }
                                    }
                                    // ������� �: ������� ����� �� ����� (�������� �����)
                                    else
                                    {
                                        // ������� ��������� ���. ������� (���� ���� ������� UseCondition)
                                        bool conditionPass = true;
                                        if (rule.UseCondition)
                                        {
                                            int condColIndex = ExcelColumnLetterToNumber(rule.ConditionColumn);
                                            if (condColIndex > 0)
                                            {
                                                // ������� � MAIN SHEET
                                                string actualValue = mainSheet.Cells[row, condColIndex].Text?.Trim();
                                                if (!string.Equals(actualValue, rule.ConditionValue, StringComparison.OrdinalIgnoreCase))
                                                    conditionPass = false;
                                            }
                                        }

                                        if (conditionPass)
                                        {
                                            string textToSearch = "";
                                            int colIndex = ExcelColumnLetterToNumber(rule.SearchColumn);

                                            // ������ ����� �� MAIN SHEET ��� ������ �������� ����
                                            if (!string.IsNullOrWhiteSpace(rule.SearchColumn) && colIndex > 0)
                                                textToSearch = mainSheet.Cells[row, colIndex].Text?.Trim();
                                            else
                                            {
                                                // ����� �� ���� ������
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
                                                        foundInMainTable = textToSearch; // ����� ����������
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // =========================================================
                                    // 2. Основной класс формы (GeneralSettingsForm)���������� ���������� (� ����������� �� ���������)
                                    // =========================================================
                                    if (matchFound)
                                    {
                                        string resultValue = "";

                                        // ��������: ����� �� �������������� �������
                                        if (rule.ResultSource == DataSourceType.AuxFile)
                                        {
                                            if (auxSheet != null)
                                            {
                                                // ����� �� �� ������ (row) � ������� ������ (SearchColumn)
                                                // (��� TargetColumn, ���� ������ ������������� ����, �� ������ ����� �� ������� "��� ������")
                                                int targetColIdx = ExcelColumnLetterToNumber(rule.SearchColumn);

                                                if (targetColIdx > 0)
                                                {
                                                    // ������ �� AUX SHEET
                                                    resultValue = auxSheet.Cells[row, targetColIdx].Text?.Trim();
                                                }
                                            }
                                            else
                                            {
                                                // ���� � ��� ����� �� ������ � ��������� ������.
                                                // �� ��������� ��� ������� �� �������� �������!
                                                // ����� �������� ���: Console.WriteLine($"���� {sheetName} �� ������ � Aux");
                                            }
                                        }
                                        // ��������: ����� �� �������� ������� (������ �����)
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(rule.VisioMasterName))
                                                resultValue = rule.VisioMasterName; // ������� ���
                                            else
                                            {
                                                // ����� �� ��������� ������ �������� �������
                                                // ���� �� ������ �� �������� "1", �� ����� ����� �� ������� SearchColumn
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

                                        // ���� ��������� ������ (��������, � ��� ������� � ���� ������ �����) � ����������
                                        if (string.IsNullOrEmpty(resultValue)) continue;

                                        // --- ����� ����������� ��������� (���������� � ����������) ---
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
                                            FullItemName = resultValue,       // <-- �������� �� ������ �������
                                            TargetMasterName = resultValue,   // <-- �������� �� ������ �������
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
                        // ��������� ���. ����, ����� �� ����� � ������
                        if (auxPackage != null) auxPackage.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingSystem.LogException("Form1", "ScanSpecificSheet", ex, 
                    new Dictionary<string, object> { { "FilePath", filePath } });
                MessageBox.Show($"������ ������������: {ex.Message}");
            }
            finally
            {
                LoggingSystem.Log(LogLevel.INFO, "Form1", "ScanSpecificSheet", 
                    $"Scan completed. Found {rawHits.Count} hits");
                MethodLogger.LogExit(rawHits);
            }

            return rawHits;
        }

        // ��������������� ����� ��� �������� ���������� (�������� ��� � ����� Form1)
        private bool TryParseQuantity(string text, out int result)
        {
            result = 1;
            if (string.IsNullOrWhiteSpace(text)) return false;

            // �������� ����� �� ������� ��� ����������� �������� � RU-������
            text = text.Replace(".", ",");

            if (double.TryParse(text, System.Globalization.NumberStyles.Any, new System.Globalization.CultureInfo("ru-RU"), out double d))
            {
                if (d > 0)
                {
                    result = (int)Math.Round(d); // ��������� (�� ������ 1.00)
                    return true;
                }
            }
            return false;
        }
        // ����������: ��������������, ��� ������� GetColumnIndex(string columnName) 
        // � ��������� RawExcelHit ��� ����������.

        // ������� ��� �������������� ��������������� ����� CountOccurrences, �� ������ �� �����
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

        private void фShowResultMessage(int totalFound)
        {

            int count = _rawHits.Count;
            string word = GetDeclension(count, "�������", "�������", "�������");
            if (count > 0)
            {
                UpdateStatus($"? ������������ ���������. ������� {count} {word}.");
            }
            else
            {
                UpdateStatus("? ������������ ���������. ������ �� �������.");
                MessageBox.Show("������ �� �������. ��������� ��������� ������.", "���������", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }



        private void GeneratePage(Visio.Document doc, string pageName, Dictionary<string, int> masterCounts, List<SearchRule> SearchRules)
        {
            // ����������: ���������� ���������� �������� SearchRules ������ ��������������� config
            var masterMap = RulesToMap(SearchRules);

            // 1. Конфигурация и настройки (основные классы)�������� ������ ����� (���� ��� ���) ��� ����� ���������
            Visio.Page page;
            try
            {
                page = doc.Pages.Add();
                page.Name = pageName;
            }
            catch (COMException)
            {
                // ���� ���� ��� ����, ���������� ���
                page = doc.Pages.get_ItemU(pageName);
                page.Background = 0;
            }

            // ��������� ����������
            double currentX = 0.5; // ��������� ������� X (� ������)
            double currentY = 0.5; // ��������� ������� Y (� ������)
            double rowHeight = 0; // ������ ������ �������� �������� � ������� ������

            // ����������: ���������������� �������������� ���������� (CS0219)
            // const double SPACING = 0.05; 
            // const double PAGE_WIDTH = 0.279; 
            const double SPACING = 0.05; // ���� �� ���������� ������������ ��, ���������������� � �������� ������
            const double PAGE_WIDTH = 0.279;

            // �������� �� ���� ���������, ������� ����� �������� (�� Excel)
            foreach (var excelKeyCount in masterCounts)
            {
                string excelKey = excelKeyCount.Key;
                int count = excelKeyCount.Value;

                // 2. Основной класс формы (GeneralSettingsForm)������� ��� ������� Visio �� ����� �� Excel
                if (masterMap.TryGetValue(excelKey, out string masterName))
                {
                    // �������� ������ �� ������ �� �������� ����������
                    Visio.Master master = doc.Masters.get_ItemU(masterName);

                    // ��������� ������
                    for (int i = 0; i < count; i++)
                    {
                        // ��������� ������ �� ��������
                        Visio.Shape shape = page.Drop(master, 0, 0); // ���������� ������� � (0,0)

                        // �������� ������� ������� ������ (������/������)
                        double shapeWidth = shape.CellsU["Width"].ResultIU;
                        double shapeHeight = shape.CellsU["Height"].ResultIU;

                        // ���������, ���������� �� ������ � ������� ������
                        if (currentX + shapeWidth > PAGE_WIDTH)
                        {
                            // ������� �� ����� ������
                            currentX = 0.5;
                            currentY += rowHeight + SPACING; // ����� ����
                            rowHeight = 0; // ���������� ������
                        }

                        // ������������� ������� ������ (�����)
                        shape.CellsU["PinX"].ResultIU = currentX + shapeWidth / 2.0;
                        shape.CellsU["PinY"].ResultIU = currentY + shapeHeight / 2.0;

                        // ��������� ��������� ��� ��������� ������
                        currentX += shapeWidth + SPACING;
                        if (shapeHeight > rowHeight)
                            rowHeight = shapeHeight;
                    }
                }
            }
            // �����������: ��������� ������ �������� ��� ����������
            page.ResizeToFitContents();
        }

        private void ShowResult(bool success, string message)
        {
            // ��������, ����� �� Invoke (���� ����� ������ �� �� UI-������)
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ShowResult(success, message)));
                return;
            }

            // ����� ���������
            if (success)
            {
                MessageBox.Show("�������� Visio ������� ������ � ������.", "�����", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"��������� ������ ��� ��������� ��������� Visio: {message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // �������� ������ �������, ���� �� �� ���������
            // btnOpenVisio.Enabled = true;
            // btnLoad.Enabled = true;
        }



        public static class VisioVbaRunner
        {
            // *** �����: ���� ����� ����������� COM-�������. �� ��������� ��� ��� doc ��� app, ���� ��� �����! ***
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

                // --- ��������� SEQUENTIAL ---
                var seq = config.SequentialDrawing;

                // ������ ��������� ����������
                double seqStartX = 10.0, seqStartY = 250.0;
                var parts = seq.StartCoordinatesXY.Split(',');
                if (parts.Length >= 2)
                {
                    double.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out seqStartX);
                    double.TryParse(parts[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out seqStartY);
                }

                // --- ��������� VBA ---
                sb.AppendLine($"Sub Draw_{moduleName}()");
                sb.AppendLine("    On Error GoTo HandleError"); // ��������� ��� ����� ����������� VBA ������
                sb.AppendLine("    Const visInches = 65");
                sb.AppendLine($"    Dim pg As Visio.Page");
                sb.AppendLine($"    Dim doc As Visio.Document");
                sb.AppendLine($"    Dim mst As Visio.Master");
                sb.AppendLine($"    Dim shp As Visio.Shape");
                sb.AppendLine($"    Dim i As Integer");
                sb.AppendLine($"    Dim mName As String"); // ��������� ��� ������ �������
                sb.AppendLine($"    Dim dropX As Double, dropY As Double");
                sb.AppendLine($"    Dim w As Double, h As Double");

                sb.AppendLine($"    Set pg = ActiveDocument.Pages.ItemU(\"{pageName}\")");
                sb.AppendLine($"    If pg Is Nothing Then GoTo CleanExit ' �������� ��������");

                // ������� ������ (�� ������ ����)
                int count = itemsToDraw.Count;
                sb.AppendLine($"    Dim masters({count}) As String");
                sb.AppendLine($"    Dim types({count}) As String");
                sb.AppendLine($"    Dim xPos({count}) As Double");
                sb.AppendLine($"    Dim yPos({count}) As Double");
                sb.AppendLine($"    Dim anchors({count}) As String");

                // ... ���������� �������� ...
                for (int j = 0; j < count; j++) // �������� i �� j, ����� �� ������������� � i � VBA �����
                {
                    var item = itemsToDraw[j];
                    // ������������ ������ ������� ��� � ����� ������� ����
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

                // ���������� ������� (Sequential)
                sb.AppendLine($"    Dim curX As Double: curX = {seqStartX.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim curY As Double: curY = {seqStartY.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim lineStart As Double: lineStart = curX");
                sb.AppendLine($"    Dim maxW As Double: maxW = {seq.MaxLineWidthMM.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim hGap As Double: hGap = {seq.HorizontalStepMM.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine($"    Dim vGap As Double: vGap = {seq.VerticalStepMM.ToString(CultureInfo.InvariantCulture)} * MM2IN");
                sb.AppendLine("    Dim rowMaxH As Double: rowMaxH = 0");

                sb.AppendLine($"    For i = 0 To {count - 1}");

                // --- �������� ������ ������ ������� --- (� ���������, ����� � ����������)
                sb.AppendLine("        Set mst = Nothing");
                sb.AppendLine("        mName = masters(i)");

                // 1. Конфигурация и настройки (основные классы)����� � ��������� (�� NameU)
                sb.AppendLine("        On Error Resume Next");
                sb.AppendLine("        Set mst = ActiveDocument.Masters.ItemU(mName)");
                sb.AppendLine("        On Error GoTo 0");

                // 2. Основной класс формы (GeneralSettingsForm)���� �� �����, ���� �� ���� �������� ����������
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
                sb.AppendLine("        Else ' ���� ������ ������, ��������� ������ ����������");
                sb.AppendLine("            Dim w As Double, h As Double");
                sb.AppendLine("            w = mst.Cells(\"Width\").Result(visInches)");
                sb.AppendLine("            h = mst.Cells(\"Height\").Result(visInches)");
                sb.AppendLine("            Dim dropX As Double, dropY As Double");

                // --- ������ ��������� (Sequential/Manual) --- (��������� ��� ���������)
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

                // �������
                sb.AppendLine("            Set shp = pg.Drop(mst, dropX, dropY)");

                // --- �������� Pin (������ ������������ ���� �������� PinY) ---
                sb.AppendLine("            If Not shp Is Nothing Then");
                sb.AppendLine("                shp.Cells(\"PinX\").ResultIU = dropX");
                sb.AppendLine("                shp.Cells(\"PinY\").ResultIU = dropY");
                sb.AppendLine("            End If");

                sb.AppendLine("        End If"); // ����� If Not mst Is Nothing Then
                sb.AppendLine("    Next i");

                sb.AppendLine("CleanExit:"); // ����� ������ ��� ���������� ������
                sb.AppendLine("    Set pg = Nothing: Set mst = Nothing: Set shp = Nothing: Set doc = Nothing"); // ������� COM-����������
                sb.AppendLine("    Exit Sub");

                sb.AppendLine("HandleError:"); // ���������� ������
                sb.AppendLine("    MsgBox \"VBA Run-time error: \" & Err.Description & \" (Code: \" & Err.Number & \") on line \" & Erl, vbCritical");
                sb.AppendLine("    Resume CleanExit");

                sb.AppendLine("End Sub");

                // ������
                Microsoft.Vbe.Interop.VBComponent? vbComp = null;
                try
                {
                    if (doc.VBProject != null)
                    {
                        vbComp = doc.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                        vbComp.Name = moduleName;
                        vbComp.CodeModule.AddFromString(sb.ToString());
                        // �����: ExecuteLine ����� ����������� ����������, ���� ���� �������� � COM (Invalid DOS Handle)
                        doc.ExecuteLine($"Draw_{moduleName}");
                    }
                }
                catch (Exception ex)
                {
                    // ����� ����� ������� ������ "������������ ���������� DOS"
                    MessageBox.Show("����������� ������ COM INTEROP: " + ex.Message +
                        "\n\n��������, ������ Visio.Application ��� Visio.Document ��� ���������� � ���������� ���� (����� Marshal.ReleaseComObject ��� ���������� using-�����) �� ����, ��� ������ �������� ������.",
                        "������ VBA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // !!! ���������� ����� ��� ������� !!!
                    if (doc.VBProject != null && vbComp != null)
                    {
                        try { doc.VBProject.VBComponents.Remove(vbComp); } catch { }
                    }
                    ReleaseComObject(vbComp);
                }
            }
        }

        // ��������������� ����� ��� �������� ������
        public class VisioItem
        {
            public string MasterName { get; set; }
            public double X { get; set; } = 0; // ���������� (� MM, �������������� � ����� � VBA Runner)
            public double Y { get; set; } = 0; // ���������� (� MM, �������������� � ����� � VBA Runner)
            public string Anchor { get; set; } = "Center";

            public string PlacementType { get; set; } = "Manual";
        }

        private async void OpenVisioClick(object? sender, EventArgs e)
        {
            _lastActivityTime = DateTime.Now; // ��������� ����������
            // ���������, ������� �� �������
            if (mainDgv == null || mainDgv.Rows.Count == 0)
            {
                MessageBox.Show("������� �������� ������� (������ '������� �������')!", "������");
                return;
            }

            this.Enabled = false;
            UpdateStatus("? ���������� ������ �� �������...");

            var tableHits = new List<RawExcelHit>();

            foreach (DataGridViewRow row in mainDgv.Rows)
            {
                if (row.IsNewRow) continue;

                // 1. Конфигурация и настройки (основные классы)������ ��� ����� (������� 3 - "BlockName" ��� ������ 2)
                var cellValue = row.Cells["BlockName"].Value?.ToString(); // ��� row.Cells[2].Value

                // 2. Основной класс формы (GeneralSettingsForm)������ ����� �/� ��� ���������� (������� 1 - "�" ��� ������ 0)
                // ���� ��� ����� ��� �� �����, ������ 0 ��� int.MaxValue (� �����)
                int sortOrder = int.MaxValue;
                var sortVal = row.Cells[0].Value?.ToString(); // ������������, ��� � � 0-� �������
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
                        SortIndex = sortOrder // ���������� �����
                    });
                }
            }

            // �����: ��������� ������ �� ����������� ������ �/� ����� ���������
            tableHits = tableHits.OrderBy(h => h.SortIndex).ToList();

            if (!tableHits.Any())
            {
                MessageBox.Show("� ������� ��� ������ � ������� '������������ �����'.");
                this.Enabled = true;
                return;
            }

            UpdateStatus("? ������ Visio...");

            await Task.Run(() =>
            {
                Visio.Application? visioApp = null;
                try
                {
                    visioApp = new Visio.Application();
                    visioApp.Visible = true;
                    var doc = visioApp.Documents.Add("");

                    // ��������� ��������� �� ���� ��������
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

                    // === ��������� 3-� ������ ===
                    // ������� GeneratePageDirectly ���� ��������:
                    // 1. Конфигурация и настройки (основные классы)���� �� ��� ����� �� ������� � �������� ����� ������� (Tab 2)
                    // 2. Основной класс формы (GeneralSettingsForm)���� ���� -> ������� MasterName
                    // 3. ������� ��������� ���������/�������� �� SequentialDrawing (Tab 3) � ���������

                    UpdateStatus("������ ����������...");
                    GeneratePageDirectly(doc, "����������", tableHits, AppSettings.LabelingConfig);

                    UpdateStatus("������ �����...");
                    GeneratePageDirectly(doc, "�����", tableHits, AppSettings.SchemeConfig);

                    UpdateStatus("������ ����...");
                    GeneratePageDirectly(doc, "����", tableHits, AppSettings.CabinetConfig);

                    UpdateStatus("? Visio ������.");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"? ������ Visio: {ex.Message}");
                    MessageBox.Show(ex.Message);
                }
            });

            this.Enabled = true;
        }

        // ������ ������ Form1

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

            // --- 1. ������������� ������ (Predefined) ---
            if (config.PredefinedMasterConfigs != null)
            {
                foreach (var fixedItem in config.PredefinedMasterConfigs)
                {
                    if (string.IsNullOrWhiteSpace(fixedItem.MasterName)) continue;

                    // ������� ���������
                    double x = 0, y = 0;
                    var coords = fixedItem.CoordinatesXY?.Split(new[] { ',', ';' });
                    if (coords != null && coords.Length >= 2)
                    {
                        double.TryParse(coords[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out x);
                        double.TryParse(coords[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out y);
                    }

                    for (int i = 0; i < fixedItem.Quantity; i++)
                    {
                        // 1. Конфигурация и настройки (основные классы)������� ������
                        Visio.Shape shp = DropShapeOnPage(page, fixedItem.MasterName, 0, 0);

                        if (shp != null)
                        {
                            // 2. Основной класс формы (GeneralSettingsForm)������������� �������
                            string anchor = !string.IsNullOrWhiteSpace(fixedItem.Anchor) ? fixedItem.Anchor : "Center";
                            SetShapePosition(shp, x, y, anchor);

                            // ====================================================================
                            // 3. �����: ���������� �����
                            // ���� ������� ��� ���� � ��������, ���������� �� � ������ ������
                            // ====================================================================
                            if (!string.IsNullOrWhiteSpace(fixedItem.FieldName) && !string.IsNullOrWhiteSpace(fixedItem.FieldValue))
                            {
                                SetShapeData(shp, fixedItem.FieldName, fixedItem.FieldValue);
                            }
                        }
                    }
                }
            }

            // --- 2. ��������� ������ (����� �� �����) ---
            // --- 2. ��������� ������ (Sequential) ---
            if (config.SearchRules != null && config.SearchRules.Any() && hits != null && hits.Any())
            {
                bool seqEnabled = config.SequentialDrawing.Enabled;

                // ��������� ���������� � ��
                double curX_MM = 10;
                double curY_MM = 200;

                // ������ ��������� ���������� �� ��������
                var sCoords = config.SequentialDrawing.StartCoordinatesXY?.Split(new[] { ',', ';' });
                if (sCoords != null && sCoords.Length >= 2)
                {
                    double.TryParse(sCoords[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out curX_MM);
                    double.TryParse(sCoords[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out curY_MM);
                }

                double maxW_MM = config.SequentialDrawing.MaxLineWidthMM;
                double hGap_MM = config.SequentialDrawing.HorizontalStepMM;
                double vGap_MM = config.SequentialDrawing.VerticalStepMM;

                // ���������� ����� �� ����������� ������ (�������� "TopLeft" ��� "Center")
                string globalAnchor = !string.IsNullOrWhiteSpace(config.SequentialDrawing.Anchor)
                                      ? config.SequentialDrawing.Anchor
                                      : "Center";

                double startX_MM = curX_MM;
                double rowMaxH_MM = 0;
                int qfCounter = 2;         // �������� (QF2...)
                int blockLabelCounter = 3; // ������ (3, 4...)
                int sensorCounter = 1;     // ������� (1, 2, 3...)

                // �������� ����� ��� ���������� ��������� �����
                string[] splitKeywords = new[] { "������.", "���������", "������.", "��������", "����", "����"};

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
                            // 1. Конфигурация и настройки (основные классы)������� �������� ������
                            Visio.Shape shp = DropShapeOnPage(page, matchedRule.VisioMasterName, 0, 0);

                            if (shp != null)
                            {
                                // --- �. ���������� ������ ---

                                // 1. Конфигурация и настройки (основные классы)������������
                                SetShapeData(shp, "������������ ������", hit.SearchTerm);

                                // 2. Основной класс формы (GeneralSettingsForm)��������� ��������� (���� ���� ���� "����� ��������")
                                if (SetShapeData(shp, "����� ��������", $"QF{qfCounter}"))
                                {
                                    qfCounter++;
                                }

                                // 3. === �����: ��������� �������� (���� ���� ���� "����� �������") ===
                                // �������� � 1. ���� �������� ������� � �����������.
                                if (SetShapeData(shp, "����� �������", sensorCounter.ToString()))
                                {
                                    sensorCounter++;
                                }

                                // --- �. ���������������� �������� ������ ---
                                string finalAnchor = !string.IsNullOrWhiteSpace(matchedRule.Anchor) ? matchedRule.Anchor : globalAnchor;

                                double mainW_MM = 0;
                                double mainH_MM = 0;
                                double mainPinX_MM = 0;
                                double mainPinY_MM = 0;

                                if (seqEnabled)
                                {
                                    // ������
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
                                    // �������
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

                                // --- �. ���������� ��������� ����� (������ ��� ����� "�����") ---
                                if (pageName.IndexOf("����", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // ���������, ����� �� ������ ���� �� 2 �����
                                    bool splitBlock = splitKeywords.Any(k => hit.SearchTerm.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);

                                    if (splitBlock)
                                    {
                                        // === ������� �: ��� ��������� ===
                                        double halfWidth = mainW_MM / 2.0;
                                        double topEdgeY = mainPinY_MM + (mainH_MM / 2.0);

                                        // ����� ���������
                                        Visio.Shape leftLbl = DropShapeOnPage(page, "0.������ ������ �������", 0, 0);
                                        if (leftLbl != null)
                                        {
                                            leftLbl.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters] = halfWidth;
                                            double lblH = leftLbl.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                            double leftPinX = mainPinX_MM - (mainW_MM / 4.0);
                                            double lblPinY = topEdgeY - (lblH / 2.0);

                                            leftLbl.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = leftPinX;
                                            leftLbl.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = lblPinY;

                                            SetShapeData(leftLbl, "��������� �����", blockLabelCounter.ToString());
                                            blockLabelCounter++;
                                        }

                                        // ������ ���������
                                        Visio.Shape rightLbl = DropShapeOnPage(page, "0.������ ������ �������", 0, 0);
                                        if (rightLbl != null)
                                        {
                                            rightLbl.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters] = halfWidth;
                                            double lblH = rightLbl.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                            double rightPinX = mainPinX_MM + (mainW_MM / 4.0);
                                            double lblPinY = topEdgeY - (lblH / 2.0);

                                            rightLbl.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = rightPinX;
                                            rightLbl.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = lblPinY;

                                            SetShapeData(rightLbl, "��������� �����", blockLabelCounter.ToString());
                                            blockLabelCounter++;
                                        }
                                    }
                                    else
                                    {
                                        // === ������� �: ���� ����� ���� ===
                                        Visio.Shape labelShp = DropShapeOnPage(page, "0.������ ������ �������", 0, 0);
                                        if (labelShp != null)
                                        {
                                            labelShp.Cells["Width"].Result[Visio.VisUnitCodes.visMillimeters] = mainW_MM;
                                            double labelH_MM = labelShp.Cells["Height"].Result[Visio.VisUnitCodes.visMillimeters];

                                            double topEdgeY = mainPinY_MM + (mainH_MM / 2.0);
                                            double newLabelPinY = topEdgeY - (labelH_MM / 2.0);

                                            labelShp.Cells["PinX"].Result[Visio.VisUnitCodes.visMillimeters] = mainPinX_MM;
                                            labelShp.Cells["PinY"].Result[Visio.VisUnitCodes.visMillimeters] = newLabelPinY;

                                            SetShapeData(labelShp, "��������� �����", blockLabelCounter.ToString());
                                            blockLabelCounter++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // �������� ������ ������ ��������
            try
            {
                // ��������� ������ ���� (���������� � Visio � 1)
                if (doc.Pages.Count > 1)
                {
                    Visio.Page firstPage = doc.Pages[1];

                    // ������� ������ ���� ��� ������ �� ����������� ("Page-1", "��������-1") � ��� ������
                    // ��� ������� ���� ����� "����������" � �.�. �� ��������
                    string name = firstPage.Name;
                    if ((name.StartsWith("Page") || name.StartsWith("��������")) && firstPage.Shapes.Count == 0)
                    {
                        firstPage.Delete(0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("������ ��� �������� ������ ��������: " + ex.Message);
            }
            // ======================================================

            UpdateStatus("? Visio ������.");
        }

        // --- ���������� ����� ������� ������ ---
        private Visio.Shape? DropShapeOnPage(Visio.Page page, string masterName, double xMM, double yMM)
        {
            // 1. Конфигурация и настройки (основные классы)�������� ������� ������
            if (string.IsNullOrWhiteSpace(masterName)) return null;
            masterName = masterName.Trim();

            Visio.Master? mst = null;
            Visio.Document doc = page.Document;
            Visio.Application app = doc.Application;

            // --- ��������� ������� ��� ������ ---
            // �������� ����� ������� �� �������������� ����� (NameU), ����� �� ���������� (Name)
            Visio.Master? TryGetMaster(Visio.Document d, string name)
            {
                // 1. Конфигурация и настройки (основные классы)���������: ������������� ��� (ItemU)
                try
                {
                    return d.Masters.get_ItemU(name);
                }
                catch { }

                // 2. Основной класс формы (GeneralSettingsForm)������: ��������� ��� ����� ���������� [ ]
                // �����: ����� ������ ������������ get_Item()
                try
                {
                    return d.Masters[name];
                }
                catch { }

                return null;
            }
            // -------------------------------------

            // 2. Основной класс формы (GeneralSettingsForm)���� � ������� ���������
            mst = TryGetMaster(doc, masterName);

            // 3. ���� �� �����, ���������� �������� ���������
            if (mst == null)
            {
                foreach (Visio.Document d in app.Documents)
                {
                    // �������� ������ � ����������� (.vssx, .vss)
                    if (d.Type == Visio.VisDocumentTypes.visTypeStencil)
                    {
                        mst = TryGetMaster(d, masterName);
                        if (mst != null) break; // ����� � ������� �� �����
                    }
                }
            }

            // 4. ���� ������ ��� � �� ������ � �����
            if (mst == null) return null;

            // 5. ������� ������ (����������� MM -> Inch)
            try
            {
                const double MM_TO_INCH = 1.0 / 25.4;
                return page.Drop(mst, xMM * MM_TO_INCH, yMM * MM_TO_INCH);
            }
            catch (Exception)
            {
                // ����������� ������ ����� �������� ����
                return null;
            }
        }

        private void SetShapePosition(Visio.Shape shp, double xMM, double yMM, string anchor)
        {
            if (shp == null) return;

            // 1. Конфигурация и настройки (основные классы)��������� ������ � ��������� ��������� "mm" ��� Visio
            // ��� ����� �������� ������ �������� �������� �����/��
            string sX = xMM.ToString(CultureInfo.InvariantCulture) + " mm";
            string sY = yMM.ToString(CultureInfo.InvariantCulture) + " mm";

            // 2. Основной класс формы (GeneralSettingsForm)����������� Anchor (����� �������� ������ ������ - LocPin)
            // FormulaU ��������� ������ "Width*0" � �.�.
            switch (anchor?.ToLower())
            {
                case "topleft":
                    shp.CellsU["LocPinX"].FormulaU = "Width*0"; // ����� ����
                    shp.CellsU["LocPinY"].FormulaU = "Height*1"; // ������� ����
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

            // 3. ������ ���� ������ � ���������� �� ����� (Pin)
            shp.CellsU["PinX"].FormulaU = sX;
            shp.CellsU["PinY"].FormulaU = sY;
        }

        // ��������������� ����� ��� ������� ����� ������
        private Visio.Shape? DropShapeOnPage(Visio.Page page, string masterName, double xMM, double yMM, int qty, bool isSequential = false)
        {
            if (string.IsNullOrWhiteSpace(masterName)) return null;

            masterName = masterName.Trim();
            Visio.Master? mst = null;
            Visio.Document doc = page.Document;
            Visio.Application app = doc.Application;

            // --- 1. ���� ������ � ����� ��������� ---

            // ������� �: �� �������������� ����� (NameU)
            try { mst = doc.Masters.get_ItemU(masterName); } catch { }

            // ������� �: �� ���������� ����� (Name) ����� ���������� [ ]
            if (mst == null)
            {
                try { mst = doc.Masters[masterName]; } catch { }
            }

            // --- 2. ���� ���, ���� �� ���� �������� ���������� ---
            if (mst == null)
            {
                foreach (Visio.Document d in app.Documents)
                {
                    // ��������� ������ ��������� (Type = 2, visTypeStencil)
                    if (d.Type == Visio.VisDocumentTypes.visTypeStencil)
                    {
                        // ������� �: NameU
                        try { mst = d.Masters.get_ItemU(masterName); } catch { }

                        // ������� �: Name ����� ���������� [ ]
                        if (mst == null)
                        {
                            try { mst = d.Masters[masterName]; } catch { }
                        }

                        if (mst != null) break; // �����!
                    }
                }
            }

            // --- 3. ���� ��� � �� ����� ---
            if (mst == null)
            {
                // ����� ����������������� ��� �������
                // Console.WriteLine($"������ '{masterName}' �� ������.");
                return null;
            }

            // --- 4. ������� ����� ---
            Visio.Shape? lastShape = null;
            const double MM_TO_INCH = 1.0 / 25.4;

            try
            {
                for (int i = 0; i < qty; i++)
                {
                    // Drop ���������� �����
                    lastShape = page.Drop(mst, xMM * MM_TO_INCH, yMM * MM_TO_INCH);
                }
            }
            catch (Exception ex)
            {
                // ����������� ������ �������, ���� �����
                Console.WriteLine($"������ ��� Drop: {ex.Message}");
            }

            return lastShape;
        }

        // ��������������� ����� ��� ������ ������
        // ����������� ���������: ��������� List<RawExcelHit> ������ Dictionary
        private void ProcessPageVba(Visio.Document doc, Visio.Page page, List<RawExcelHit> hits, VisioConfiguration config)
        {
            var itemsToDraw = new List<VisioItem>();

            // 1. Конфигурация и настройки (основные классы)������������� ������ (Manual) - �������� ��� ���������
            if (config.PredefinedMasterConfigs != null)
            {
                foreach (var pm in config.PredefinedMasterConfigs)
                {
                    if (string.IsNullOrWhiteSpace(pm.MasterName)) continue;

                    // ������� ��������� � ���������� ����� � �������
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

            // 2. Основной класс формы (GeneralSettingsForm)��������� ������ (Sequential)
            if (hits != null && config.SequentialDrawing.Enabled)
            {
                // ���������� ���������� ������, ����� �� �������� ���������, ���� ��� �� �����
                // ���������� �� SearchTerm, ��� ��� ������ �� �������� � MasterName � �������
                var groupedHits = hits
                    .Where(h => h.ConditionMet) // ������ ��, ��� ������� ���������
                    .GroupBy(h => h.SearchTerm);

                foreach (var group in groupedHits)
                {
                    string searchTerm = group.Key;

                    // ������� ������� � ������� ��� ����� ���������� �����
                    var rule = config.SearchRules.FirstOrDefault(r =>
                        string.Equals(r.ExcelValue, searchTerm, StringComparison.OrdinalIgnoreCase));

                    // ���� ������� ������� � � ���� ���� ��� ������� Visio
                    if (rule != null && !string.IsNullOrWhiteSpace(rule.VisioMasterName))
                    {
                        // ������� ����������
                        int totalQty = 0;

                        // ��������� ������ ��������� � ������
                        foreach (var hit in group)
                        {
                            if (hit.IsLimited)
                                totalQty += 1; // ���� ���������� - ������� ��� 1 (�� ��� �����: ����������� ������ �� ������)
                            else
                                totalQty += hit.Quantity;
                        }

                        // ��������� ������ LIMIT: 
                        // ���� � ������� ����� LimitQuantity, �� �� ������ ���������� ������ ������ ���� ���,
                        // ���������� �� ����, ������� ����� �� �����.
                        if (rule.LimitQuantity)
                        {
                            totalQty = 1;
                        }

                        for (int k = 0; k < totalQty; k++)
                        {
                            itemsToDraw.Add(new VisioItem
                            {
                                MasterName = rule.VisioMasterName.Trim(),
                                X = 0, // ������������ ��� Sequential
                                Y = 0,
                                PlacementType = "Sequential",
                                Anchor = config.SequentialDrawing.Anchor
                            });
                        }
                    }
                }
            }

            // ������ �������
            VisioVbaRunner.RunDrawingMacro(doc, page, itemsToDraw, config);
        }

        /// <summary>
        /// ��������� ����������� COM-������, ������� Marshal.ReleaseComObject � �����, 
        /// ���� ������� ������ �� ������ ������ ����.
        /// ��� ���������� ����� ��� ����������� ���������� ������ � Visio Interop.
        /// </summary>
        private void ReleaseComObject(object? obj)
        {
            // ���������, ��� ������ ���������� � �������� COM-��������
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {
                    // ��������� ���� ��� ���������������� ������������.
                    // ���� ������ ��� ������� ���������� ����������, Marshal.ReleaseComObject
                    // ���������� ������� ������, ������� ������ ����� 0 ��� ������� ������������.
                    while (Marshal.ReleaseComObject(obj) > 0)
                    {
                        Marshal.ReleaseComObject(obj);
                    }
                }
                catch (Exception ex)
                {
                    // ����� ����� �������� ������������, ���� �� ������� ���������� ������
                    System.Diagnostics.Debug.WriteLine($"������ ��� ������������ COM-�������: {ex.Message}");
                }
                finally
                {
                    // �������� ������ � ����������� ����
                    obj = null;
                }
            }
        }

        /// <summary>
        /// ������������ "������������" �� Excel � ������ Visio Master �� MasterMap.
        /// </summary>
        private List<Dictionary<string, string>> PrepareVisioData(
    List<Dictionary<string, string>> extractedData,
    List<SearchRule> SearchRules)
        {
            // ����������: ���������� �������� SearchRules ������ ��������������� config.SearchRules
            var masterMap = RulesToMap(SearchRules);

            var visioData = new List<Dictionary<string, string>>();
            int totalItems = extractedData.Count;
            int mappedItems = 0;

            if (!masterMap.Any())
            {
                UpdateStatus("?? MasterMap ����! ���������� ����������� ������.");
                return visioData;
            }

            UpdateStatus($"������ ������������� {totalItems} ������� Excel � MasterMap...");

            foreach (var item in extractedData)
            {
                if (!item.TryGetValue("������������", out string? content) || string.IsNullOrEmpty(content))
                    continue;

                string cleanedContent = content.Trim();
                bool matched = false;

                // ���� ����� ������� ����������
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
                    UpdateStatus($"  ? ������������: '{bestMatch}' -> '{visioMasterName}'");
                }

                if (!matched)
                {
                    UpdateStatus($"  ? �� ������������: '{cleanedContent}'");
                }
            }

            UpdateStatus($"? ������������� ���������. ������� {mappedItems} �� {totalItems} �������.");
            return visioData;
        }

        // ���� ����� �������� ������ GenerateVisioDocument � ��������� ������
        // �������� ������� ��������� � ����� ���������� � �� ��������� Visio.
        // ����������� CS1501: ��������� ������ ������ ��������� 3 ���������, ��� � OpenVisioClick
        // ����������� CS1501: ��������� ������ ������ ��������� 3 ���������, ��� � OpenVisioClick

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
                UpdateStatus("?? ������ ��������� ������������� Visio-�����...");

                // ����������: 
                // 1. Конфигурация и настройки (основные классы)������ ������ ������ RulesToMap, ��� ��� PrepareVisioData ������ ��� ������.
                // 2. Основной класс формы (GeneralSettingsForm)�������� extractedData ������ ����������.
                // 3. �������� ���������� ������ ������ (configMarking.SearchRules) ������ ����������.

                var markingData = PrepareVisioData(extractedData, configMarking.SearchRules);
                var schemeData = PrepareVisioData(extractedData, configScheme.SearchRules);

                // 2. Основной класс формы (GeneralSettingsForm)������ VISIO � �������� ���������
                visioApp = new Visio.Application();
                visioApp.Visible = true; // ��������� Visio ��������

                newDocument = visioApp.Documents.Add(""); // ������� ����� ��������

                // 3. �������� �������
                Visio.Page pageMarking = newDocument.Pages[1];
                pageMarking.Name = "����������";
                pageMarking.PageSheet.CellsU["PageUnits"].FormulaU = "8"; // METER(1)
                pageMarking.PageSheet.CellsU["DrawingUnits"].FormulaU = "8"; // METER(1)

                Visio.Page pageScheme = newDocument.Pages.Add();
                pageScheme.Name = "�����";
                pageScheme.PageSheet.CellsU["PageUnits"].FormulaU = "8";
                pageScheme.PageSheet.CellsU["DrawingUnits"].FormulaU = "8";

                // 4. ���������� �������
                PopulateVisioPage(pageMarking, markingData, configMarking, false);
                PopulateVisioPage(pageScheme, schemeData, configScheme, true);

                // 5. ���������� ����� ��� ���������
                if (visioApp.ActiveWindow != null)
                {
                    visioApp.ActiveWindow.Page = pageScheme;
                }

                UpdateStatus($"? ���� Visio ������� ������������ � ������");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                UpdateStatus($"? COM ������ Visio: {ex.Message}");
                MessageBox.Show($"COM ������ Visio: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                UpdateStatus($"? ����� ������ ��� ��������� Visio: {ex.Message}");
                MessageBox.Show($"����� ������ ��� ��������� Visio: {ex.Message}", "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    // ��������� �������� �������
                    stencilDoc = stencils.OpenEx(stencilPath, (short)Visio.VisOpenSaveArgs.visOpenHidden);

                    // �������� ���� ��������
                    foreach (Visio.Master master in stencilDoc.Masters)
                    {
                        // �����������: ��������� ������� � ������ ��������
                        if (!allFoundMasters.Contains(master.Name, StringComparer.OrdinalIgnoreCase))
                        {
                            allFoundMasters.Add(master.Name);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"������ ��� �������� ��������� {stencilPath}: {ex.Message}");
                }
                finally
                {
                    // ����������� COM-������ ���������
                    if (stencilDoc != null) Marshal.ReleaseComObject(stencilDoc);
                }
            }
            if (stencils != null) Marshal.ReleaseComObject(stencils);
        }

        // ����� ��������������� �����, ����������� ������ ���������� �����
        private void PopulateVisioPage(Visio.Page page, List<Dictionary<string, string>> extractedData,
                              VisioConfiguration config, bool isScheme)
        {
            // ����������: SPACING ������ ������������ � ��, � ����� ��������������
            const double SPACING_MM = 50.0; // 50 �� ����� �������� (5 ��)
            const double INITIAL_OFFSET_MM = 25.4;
            double SPACING_INCH = SPACING_MM * MM_TO_INCH; // ������������ � �����

            // ������� const double PAGE_WIDTH = 0.297; - ��� �� ������������ ��� ��������� ����� �����

            // ����������: ��������� ���������� � ������
            double initialOffsetInches = INITIAL_OFFSET_MM * MM_TO_INCH; // 1 ���� �� ����
            double currentX = initialOffsetInches;
            double currentY = initialOffsetInches;

            var openStencils = new List<Visio.Document>();
            var mastersNotFound = new List<string>(); // ��� ����� ����������� ��������
            var allAvailableMasterNames = new HashSet<string>(StringComparer.Ordinal); // ��� ����� ���� ��������� ����

            Visio.Master? master = null;
            Visio.Shape? shape = null;

            try
            {
                UpdateStatus($"������ ���������� ����� �� �������� '{page.Name}'. ������� ������� ���������...");

                // 1. Конфигурация и настройки (основные классы)�������� ���� ���������� (��� ���������)
                foreach (string path in config.StencilFilePaths)
                {
                    if (string.IsNullOrEmpty(path)) continue;
                    try
                    {
                        if (!System.IO.File.Exists(path))
                        {
                            UpdateStatus($"? ������: ���� ��������� �� ����������: {path}");
                            continue;
                        }

                        Visio.Document stencilDoc = page.Application.Documents.Open(path);
                        openStencils.Add(stencilDoc);
                        UpdateStatus($"? �������� ������: {Path.GetFileName(path)}");
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"? ����������� ������ �������� ��������� '{Path.GetFileName(path)}': {ex.Message}");
                    }
                }

                if (!openStencils.Any())
                {
                    UpdateStatus("? ����������� ������: �� ������� ������� �� ���� �������� Visio. ����������.");
                    return;
                }

                // 2. Основной класс формы (GeneralSettingsForm)���� ���� ��������� �������� (��� �����������)
                foreach (var stencilDoc in openStencils)
                {
                    foreach (Visio.Master m in stencilDoc.Masters)
                    {
                        // �������� ������ ����� (NameU)
                        allAvailableMasterNames.Add(m.NameU);
                        ReleaseComObject(m); // ����������� ������ ����������
                    }


                }

                UpdateStatus($"������� {allAvailableMasterNames.Count} ���������� ����� � �������� ����������.");

                // =========================================================================
                // 3. ���������� ���������������� ����� (����� ������)
                // =========================================================================
                // ... ������ PopulateVisioPage ...

                // =========================================================================
                // 3. ���������� ���������������� ����� (����� ������)
                // =========================================================================
                UpdateStatus($"���������� {config.PredefinedMasterConfigs.Count} ���������������� �����...");

                // ����������: ��� ���������� ����� ������� �� var (PredefinedMasterConfig)
                foreach (var pmConfig in config.PredefinedMasterConfigs.Where(n => !string.IsNullOrWhiteSpace(n.MasterName)))
                {
                    master = null;
                    shape = null;

                    // ����������: ����� ��� �� ������� �������
                    string predefinedMasterName = pmConfig.MasterName;

                    // 3.1. ����� �������
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
                            // ������ ����������, ���� ������ �� ������ � ���� ���������
                        }
                    }

                    if (master != null)
                    {
                        try
                        {
                            // 3.2. ���������� ������: �������� �����������: ���������� ���������� �� �������
                            double xDrop = currentX, yDrop = currentY; // Fallback � ��������������� ����������
                            bool usedAutoPlacement = true;

                            // ���������, ������ �� ���������� � �������
                            if (!string.IsNullOrWhiteSpace(pmConfig.CoordinatesXY))
                            {
                                var parts = pmConfig.CoordinatesXY.Split(',');
                                if (parts.Length >= 2 &&
                                    float.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float xMM) &&
                                    float.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float yMM))
                                {
                                    // ����������� �� MM � INCH
                                    xDrop = xMM * MM_TO_INCH;
                                    yDrop = yMM * MM_TO_INCH;
                                    usedAutoPlacement = false;
                                }
                            }

                            shape = page.Drop(master, xDrop, yDrop); // ���������� ���������������� ��� �������������� ����������
                            shape.Text = $"����������������: {predefinedMasterName}";

                            // ������ ��������������� �������� (������ ���� �� �������������� ���������� �� �������)
                            if (usedAutoPlacement)
                            {
                                currentX += SPACING_INCH;
                                if (currentX > 10.0)
                                {
                                    currentX = initialOffsetInches;
                                    currentY += SPACING_INCH;
                                }
                            }

                            UpdateStatus($"  ? ��������� ���������������� ������: '{predefinedMasterName}'");
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"? ������ ���������� ����������������� ������� '{predefinedMasterName}': {ex.Message}");
                        }
                    }
                    else
                    {
                        // ���� ���������������� ������ �� ������
                        if (!mastersNotFound.Contains(predefinedMasterName))
                        {
                            mastersNotFound.Add(predefinedMasterName);
                        }
                    }

                    // ����������� �������, ��������� � ���� �����
                    ReleaseComObject(shape);
                    ReleaseComObject(master);
                }

                // 3. ���������� ����� �� ��������
                foreach (var dataItem in extractedData)
                {
                    if (!dataItem.ContainsKey("VisioMasterName") ||
                        !int.TryParse(dataItem.GetValueOrDefault("����������", "0"), out int count)) continue;

                    string masterName = dataItem["VisioMasterName"];
                    master = null;

                    // 3.1. ������� �������� (��� �� �����������, ��� Visio ������ ������, �� ���� ��� ������ �����������)
                    if (!allAvailableMasterNames.Contains(masterName))
                    {
                        if (!mastersNotFound.Contains(masterName)) mastersNotFound.Add(masterName);
                        UpdateStatus($"? ������ '{masterName}' ����������� � ������ ��������� �����.");
                        continue;
                    }

                    // 3.2. ����� �� ����������� ��������� (���� ��� ������� ����������)
                    foreach (var stencilDoc in openStencils)
                    {
                        // ��� ������� ������ ����� ��������� � NameU
                        master = stencilDoc.Masters.Cast<Visio.Master>().FirstOrDefault(m =>
                            m.NameU.Equals(masterName, StringComparison.Ordinal));

                        if (master != null)
                        {
                            break;
                        }
                    }

                    // 3.3. ��������� ������, ����� ������ �� ������ (���� ��� ���� � ������)
                    if (master == null)
                    {
                        // ��� ������������� ������ (������ COM-������ �� ����� ���� �������)
                        if (!mastersNotFound.Contains(masterName)) mastersNotFound.Add(masterName);
                        UpdateStatus($"? �� ������� �������� COM-������ ������� '{masterName}', ���� �� ��� ������ � ������.");
                        continue;
                    }

                    // 3.4. ���������� �����
                    try
                    {
                        for (int i = 0; i < count; i++)
                        {
                            shape = page.Drop(master, currentX, currentY); // currentX/Y ������ � ������
                            shape.Text = masterName;

                            // ������ ��������
                            currentX += SPACING_INCH; // ����������: ���������� SPACING_INCH
                            if (currentX > 10.0)
                            {
                                currentX = initialOffsetInches;
                                currentY += SPACING_INCH; // ����������: ���������� SPACING_INCH
                            }

                            ReleaseComObject(shape);
                            shape = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"? ������ ���������� ������� '{masterName}': {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObject(master);
                    }
                }

                UpdateStatus($"? ���������� ����� �� �������� '{page.Name}' ���������.");

                // 4. ����� ��������� � ��������� ��������
                if (mastersNotFound.Any())
                {
                    string missingMastersList = string.Join(Environment.NewLine, mastersNotFound.Distinct());

                    // ������������ ������ ��������� ���� ��� �������� ������
                    string availableList = allAvailableMasterNames.Any()
                        ? string.Join(", ", allAvailableMasterNames.OrderBy(n => n).Take(50))
                        : "��� ��������� �����.";

                    MessageBox.Show(
                        this,
                        "? ����������� ������: �� ������� ���������� ��������� ������ (�������):\n\n" +
                        $"������� ������:\n{missingMastersList}\n\n" +
                        "==========================================================\n" +
                        "��������� ������ � ���������� (Master.NameU):\n" +
                        $"{availableList}" +
                        (allAvailableMasterNames.Count > 50 ? $"\n... ����� {allAvailableMasterNames.Count} �����." : "") +
                        "\n\n����������: ������� ����� ������ ����� ��������� � ���������� ������� (������� �������!).",
                        $"�������� � �������� �� �������� '{page.Name}'",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            finally
            {
                // 5. ������� COM-��������
                foreach (var stencilDoc in openStencils)
                {
                    try
                    {
                        stencilDoc.Close();
                    }
                    catch { /* ���������� ������ ��� �������� */ }
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

        // --- �������� ����� ������ ������ FORM1 ---

        // 1. Конфигурация и настройки (основные классы)������� ������ Excel � ����������� ��� � ������ ��������
        private List<EquipmentItem> ParseExcelData(ExcelWorksheet sheet)
        {
            var foundEquipment = new List<EquipmentItem>();

            // �����: ��������� ������ ����� � ������� ��� ��� ����!
            int startRow = 10; // � ����� ������ ���������� ������
            int endRow = sheet.Dimension?.End.Row ?? 100;

            for (int row = startRow; row <= endRow; row++)
            {
                // �����������, ������������ � 1-� �������, ���-�� � 5-� (��������� �������!)
                string excelName = sheet.Cells[row, 1].Text.Trim();
                string qtyText = sheet.Cells[row, 5].Text.Trim();

                if (string.IsNullOrEmpty(excelName)) continue;

                // ���� ���������� � ����� ������� ��������
                if (_equipmentConfig.TryGetValue(excelName, out EquipmentItem configItem))
                {
                    int.TryParse(qtyText, out int qty);
                    if (qty > 0)
                    {
                        // ������� ����� ������� � �������� �����������
                        foundEquipment.Add(new EquipmentItem
                        {
                            OriginalName = excelName,
                            ShortName = configItem.ShortName,
                            // ����������: ������� ��������� ������ Count(...), � ����� ���������� + 1
                            PositionCode = configItem.PositionCode + (foundEquipment.Count(x => x.PositionCode != null && x.PositionCode.StartsWith(configItem.PositionCode)) + 1),
                            ShapeMasterName = configItem.ShapeMasterName,
                            Quantity = qty
                        });
                    }
                }
            }
            return foundEquipment;
        }

        // 2. Основной класс формы (GeneralSettingsForm)������� ������ ������ ������ ������ Visio (Shape Data)
        private void UpdateVisioShapeData(Visio.Shape shape, EquipmentItem item)
        {
            try
            {
                // ����� ����� �� ������
                shape.Text = $"{item.PositionCode}\n{item.ShortName}";

                // ��������� ������� ������ �������
                if (shape.get_SectionExists((short)Visio.VisSectionIndices.visSectionProp, 0) == 0)
                    shape.AddSection((short)Visio.VisSectionIndices.visSectionProp);

                // ���������� ������� ������ (��� ������� Visio)
                SetShapeProperty(shape, "ShortName", "������������", item.ShortName);
                SetShapeProperty(shape, "Position", "�������", item.PositionCode);
                SetShapeProperty(shape, "Quantity", "���-��", item.Quantity.ToString());
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("������ ���������� ������: " + ex.Message); }
        }

        // 3. ��������������� ������� ��� �������
        private void SetShapeProperty(Visio.Shape shape, string propName, string label, string value)
        {
            string cellName = "Prop." + propName;
            if (shape.get_CellExists(cellName, (short)Visio.VisExistsFlags.visExistsAnywhere) == 0)
                shape.AddNamedRow((short)Visio.VisSectionIndices.visSectionProp, propName, (short)Visio.VisRowTags.visTagDefault);

            shape.CellsU[cellName].FormulaU = "\"" + value + "\"";
            shape.CellsU[cellName + ".Label"].FormulaU = "\"" + label + "\"";
        }

        // 4. ������� ��������� �������
        private void DrawSpecificationTable(Visio.Page page, List<EquipmentItem> items)
        {
            double x = 1.0; // ������ ����� (�����)
            double y = 10.0; // ������ ������ (�����, Visio ������� �� ������� ����, ������� 10 - ��� ������)
            double rowHeight = 0.25;

            // ���������
            page.DrawRectangle(x, y, x + 0.5, y + rowHeight).Text = "���.";
            page.DrawRectangle(x + 0.5, y, x + 2.5, y + rowHeight).Text = "Отменить";
            page.DrawRectangle(x + 2.5, y, x + 3.0, y + rowHeight).Text = "���.";

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
    // 5. ����� ����� ������ ������ (SheetSelectionForm)
    // =========================================================================

    public class SheetSelectionForm : Form
    {
        private readonly CheckedListBox _clbSheets;
        public List<string> SelectedSheets { get; private set; } = new List<string> { "1.�� �� ������ ZONT ��� VISIO" };

        public SheetSelectionForm(List<string> allSheetNames, List<string> initialSelectedSheets)
        {
            this.Text = "����� ������ ��� �������";
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
                SelectionMode = SelectionMode.One // <-- ����������: ��������� SelectionMode � One ��� ���������� ������ ������ �� ���������.
            };
            mainLayout.Controls.Add(_clbSheets, 0, 0);

            // ��������� ������ � ������������� ��������� �������
            foreach (var sheetName in allSheetNames)
            {
                // ���������, ��� �� ���� ���� ������ ����� (��� ����� ��������)
                bool isChecked = initialSelectedSheets.Contains(sheetName, StringComparer.OrdinalIgnoreCase);
                _clbSheets.Items.Add(sheetName, isChecked);
            }

            // ������
            var footerFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(5)
            };
            mainLayout.Controls.Add(footerFlow, 0, 1);

            var btnOk = new Button { Text = "Отменить", Width = 100, Height = 30, DialogResult = DialogResult.OK };
            btnOk.Click += (s, e) =>
            {
                // �������� ��������� �����
                SelectedSheets = _clbSheets.CheckedItems.Cast<string>().ToList();
            };

            var btnCancel = new Button { Text = "Отменить", Width = 100, Height = 30, DialogResult = DialogResult.Cancel };

            footerFlow.Controls.Add(btnCancel);
            footerFlow.Controls.Add(btnOk);
        }
    }

}



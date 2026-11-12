using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using Syncfusion.XlsIO;
using EFM_BOM_App.Services;

namespace EFM_BOM_App
{
    public partial class DataTransformationWindow : Window
    {
        private readonly IWorkbook _workbook;
        private DataTable _currentTable;
        private DataTable _initialTable;
        private readonly List<string> _appliedSteps = new();

        public DataTransformationWindow(IWorkbook workbook, IWorksheet? sourceSheet = null)
        {
            InitializeComponent();
            _workbook = workbook;
            _currentTable = new DataTable();
            _initialTable = new DataTable();

            if (sourceSheet != null && sourceSheet.UsedRange != null && sourceSheet.UsedRange.Rows.Length > 0)
            {
                LoadFromWorksheet(sourceSheet);
            }
            else
            {
                CreateDemoData();
            }
        }

        private void CreateDemoData()
        {
            var dt = new DataTable("SourceData");
            dt.Columns.Add("TransactionID", typeof(string));
            dt.Columns.Add("Customer_Region", typeof(string));
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Amount", typeof(double));
            dt.Columns.Add("Status", typeof(string));

            dt.Rows.Add("TX1001", "North_Retail", "Electronics", 450.00, "Completed");
            dt.Rows.Add("TX1002", "South_Enterprise", "Office", 120.50, "Pending");
            dt.Rows.Add("TX1003", "West_Retail", "Electronics", 890.00, "Completed");
            dt.Rows.Add("TX1004", "North_Retail", "Electronics", 450.00, "Completed"); // Duplicate
            dt.Rows.Add("TX1005", "East_Enterprise", "Services", 1500.00, "Completed");
            dt.Rows.Add("TX1006", "South_Retail", "Office", 310.25, "Cancelled");
            dt.Rows.Add("TX1007", "West_Enterprise", "Services", 2100.00, "Completed");

            SetTable(dt, "Source (Demo Transactions)");
        }

        private void LoadFromWorksheet(IWorksheet sheet)
        {
            try
            {
                var dt = new DataTable(sheet.Name);
                var range = sheet.UsedRange;
                int rows = range.Rows.Length;
                int cols = range.Columns.Length;

                if (rows == 0 || cols == 0)
                {
                    CreateDemoData();
                    return;
                }

                for (int c = 1; c <= cols; c++)
                {
                    string header = range[1, c].Text;
                    if (string.IsNullOrWhiteSpace(header)) header = $"Column_{c}";
                    dt.Columns.Add(header, typeof(string));
                }

                for (int r = 2; r <= rows; r++)
                {
                    var rowVals = new object[cols];
                    for (int c = 1; c <= cols; c++)
                    {
                        rowVals[c - 1] = range[r, c].Text;
                    }
                    dt.Rows.Add(rowVals);
                }

                SetTable(dt, $"Source ({sheet.Name})");
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Failed to load sheet into Data Transformation", ex);
                CreateDemoData();
            }
        }

        private void SetTable(DataTable dt, string stepDescription)
        {
            _currentTable = dt;
            if (_appliedSteps.Count == 0)
            {
                _initialTable = dt.Copy();
            }
            _appliedSteps.Add(stepDescription);
            RefreshView();
        }

        private void RefreshView()
        {
            GridPreview.ItemsSource = null;
            GridPreview.ItemsSource = _currentTable.DefaultView;
            ListSteps.ItemsSource = null;
            ListSteps.ItemsSource = _appliedSteps;
            TxtRowCount.Text = $"Rows: {_currentTable.Rows.Count} | Columns: {_currentTable.Columns.Count}";
            TxtStatus.Text = $"Step applied: {_appliedSteps.LastOrDefault() ?? "Ready"}";
        }

        private void BtnLoadSource_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Delimited/Excel Files (*.csv;*.txt;*.xlsx)|*.csv;*.txt;*.xlsx|All Files (*.*)|*.*",
                Title = "Import Source Dataset"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    string ext = Path.GetExtension(dlg.FileName).ToLower();
                    var dt = new DataTable("Imported");

                    if (ext == ".csv" || ext == ".txt")
                    {
                        string[] lines = File.ReadAllLines(dlg.FileName, Encoding.UTF8);
                        if (lines.Length == 0) return;
                        char delimiter = lines[0].Contains('\t') ? '\t' : ',';
                        string[] headers = lines[0].Split(delimiter);
                        foreach (var h in headers) dt.Columns.Add(h.Trim('\"', ' '));

                        for (int i = 1; i < lines.Length; i++)
                        {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            string[] parts = lines[i].Split(delimiter);
                            var rowVals = new object[headers.Length];
                            for (int j = 0; j < Math.Min(parts.Length, headers.Length); j++)
                                rowVals[j] = parts[j].Trim('\"', ' ');
                            dt.Rows.Add(rowVals);
                        }
                    }
                    else
                    {
                        using (var eng = new ExcelEngine())
                        {
                            var wb = eng.Excel.Workbooks.Open(dlg.FileName);
                            var ws = wb.Worksheets[0];
                            LoadFromWorksheet(ws);
                            return;
                        }
                    }

                    _appliedSteps.Clear();
                    SetTable(dt, $"Imported ({Path.GetFileName(dlg.FileName)})");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"File import failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnRemoveDuplicates_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var uniqueTable = _currentTable.Clone();
                var seen = new HashSet<string>();

                foreach (DataRow row in _currentTable.Rows)
                {
                    string key = string.Join("|||", row.ItemArray.Select(x => x?.ToString() ?? ""));
                    if (seen.Add(key))
                    {
                        uniqueTable.ImportRow(row);
                    }
                }

                int removed = _currentTable.Rows.Count - uniqueTable.Rows.Count;
                _currentTable = uniqueTable;
                _appliedSteps.Add($"Remove Duplicates ({removed} removed)");
                RefreshView();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Remove Duplicates failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnFilterRows_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count == 0) return;

            string colName = _currentTable.Columns[0].ColumnName;
            var prompt = new SimpleInputDialog("Filter Rows", $"Keep rows where column '{colName}' does NOT equal:", "Cancelled");
            if (prompt.ShowDialog() == true)
            {
                string excludeVal = prompt.InputValue.Trim();
                var filtered = _currentTable.Clone();
                foreach (DataRow row in _currentTable.Rows)
                {
                    if (row[colName]?.ToString() != excludeVal)
                    {
                        filtered.ImportRow(row);
                    }
                }
                _currentTable = filtered;
                _appliedSteps.Add($"Filter Rows ({colName} != '{excludeVal}')");
                RefreshView();
            }
        }

        private void BtnSplitColumn_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count < 2) return;
            string targetCol = _currentTable.Columns[1].ColumnName;

            var prompt = new SimpleInputDialog("Split Column", $"Split column '{targetCol}' by delimiter:", "_");
            if (prompt.ShowDialog() == true)
            {
                string delim = prompt.InputValue;
                var splitTable = _currentTable.Copy();
                string newCol1 = $"{targetCol}_1";
                string newCol2 = $"{targetCol}_2";

                splitTable.Columns.Add(newCol1, typeof(string));
                splitTable.Columns.Add(newCol2, typeof(string));

                foreach (DataRow row in splitTable.Rows)
                {
                    string orig = row[targetCol]?.ToString() ?? "";
                    string[] parts = orig.Split(new[] { delim }, 2, StringSplitOptions.None);
                    row[newCol1] = parts.Length > 0 ? parts[0] : "";
                    row[newCol2] = parts.Length > 1 ? parts[1] : "";
                }

                splitTable.Columns.Remove(targetCol);
                _currentTable = splitTable;
                _appliedSteps.Add($"Split Column ('{targetCol}' by '{delim}')");
                RefreshView();
            }
        }

        private void BtnMergeColumns_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count < 2) return;
            string col1 = _currentTable.Columns[0].ColumnName;
            string col2 = _currentTable.Columns[1].ColumnName;

            var prompt = new SimpleInputDialog("Merge Columns", $"Merge '{col1}' and '{col2}' with separator:", " - ");
            if (prompt.ShowDialog() == true)
            {
                string sep = prompt.InputValue;
                var merged = _currentTable.Copy();
                string newCol = $"{col1}_{col2}";
                merged.Columns.Add(newCol, typeof(string));

                foreach (DataRow row in merged.Rows)
                {
                    row[newCol] = $"{row[col1]}{sep}{row[col2]}";
                }

                _currentTable = merged;
                _appliedSteps.Add($"Merge Columns ('{col1}' & '{col2}')");
                RefreshView();
            }
        }

        private void BtnReplaceValues_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count == 0) return;
            var prompt = new SimpleInputDialog("Replace Values", "Find text to replace:", "Pending");
            if (prompt.ShowDialog() == true)
            {
                string find = prompt.InputValue;
                var prompt2 = new SimpleInputDialog("Replace Values", $"Replace '{find}' with:", "In Review");
                if (prompt2.ShowDialog() == true)
                {
                    string repl = prompt2.InputValue;
                    var table = _currentTable.Copy();
                    foreach (DataRow row in table.Rows)
                    {
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            if (row[i]?.ToString() == find) row[i] = repl;
                        }
                    }
                    _currentTable = table;
                    _appliedSteps.Add($"Replace Values ('{find}' ➔ '{repl}')");
                    RefreshView();
                }
            }
        }

        private void BtnGroupBy_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count < 2) return;

            string groupCol = _currentTable.Columns[1].ColumnName;
            var dtGroup = new DataTable($"Grouped_{groupCol}");
            dtGroup.Columns.Add(groupCol, typeof(string));
            dtGroup.Columns.Add("Record_Count", typeof(int));

            var groups = _currentTable.AsEnumerable()
                .GroupBy(r => r[groupCol]?.ToString() ?? "Unknown");

            foreach (var g in groups)
            {
                dtGroup.Rows.Add(g.Key, g.Count());
            }

            _currentTable = dtGroup;
            _appliedSteps.Add($"Group & Aggregate (by '{groupCol}')");
            RefreshView();
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            _currentTable = _initialTable.Copy();
            _appliedSteps.Clear();
            _appliedSteps.Add("Reset to Source");
            RefreshView();
        }

        private void BtnApplyAndLoad_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string sheetName = "Transformed_" + DateTime.Now.ToString("HHmmss");
                var sheet = _workbook.Worksheets.Create(sheetName);

                // Write Headers
                for (int c = 0; c < _currentTable.Columns.Count; c++)
                {
                    sheet.Range[1, c + 1].Text = _currentTable.Columns[c].ColumnName;
                    sheet.Range[1, c + 1].CellStyle.Font.Bold = true;
                }

                // Write Rows
                for (int r = 0; r < _currentTable.Rows.Count; r++)
                {
                    for (int c = 0; c < _currentTable.Columns.Count; c++)
                    {
                        var val = _currentTable.Rows[r][c];
                        if (val is double d) sheet.Range[r + 2, c + 1].Number = d;
                        else if (val is int i) sheet.Range[r + 2, c + 1].Number = i;
                        else sheet.Range[r + 2, c + 1].Text = val?.ToString() ?? "";
                    }
                }

                sheet.UsedRange.AutofitColumns();
                LoggingService.LogInfo($"Transformed dataset applied and loaded into worksheet '{sheetName}' ({_currentTable.Rows.Count} rows).");
                MessageBox.Show($"Transformed dataset successfully loaded into sheet '{sheetName}' ({_currentTable.Rows.Count} rows, {_currentTable.Columns.Count} columns).", "Transformation Applied", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Failed to apply and load transformed data", ex);
                MessageBox.Show($"Load Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}


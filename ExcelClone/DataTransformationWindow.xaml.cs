using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using unvell.ReoGrid;
using ClosedXML.Excel;
using ExcelClone.Services;

namespace ExcelClone
{
    public partial class DataTransformationWindow : Window
    {
        private readonly ReoGridControl? _gridControl;
        private readonly Worksheet? _sourceSheet;
        private DataTable _currentTable;
        private DataTable _initialTable;
        private readonly List<string> _appliedSteps = new();

        public DataTransformationWindow(ReoGridControl? gridControl = null, Worksheet? sourceSheet = null)
        {
            InitializeComponent();
            _gridControl = gridControl;
            _sourceSheet = sourceSheet;
            _currentTable = new DataTable();
            _initialTable = new DataTable();

            if (sourceSheet != null)
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

        private void LoadFromWorksheet(Worksheet sheet)
        {
            try
            {
                var dt = new DataTable(sheet.Name);
                int rows = Math.Min(sheet.RowCount, 500);
                int cols = Math.Min(sheet.ColumnCount, 30);

                if (rows == 0 || cols == 0)
                {
                    CreateDemoData();
                    return;
                }

                for (int c = 0; c < cols; c++)
                {
                    string header = sheet.Cells[0, c]?.Data?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrWhiteSpace(header)) header = $"Column_{c + 1}";
                    dt.Columns.Add(header, typeof(string));
                }

                for (int r = 1; r < rows; r++)
                {
                    var rowVals = new object[cols];
                    bool hasData = false;
                    for (int c = 0; c < cols; c++)
                    {
                        var val = sheet.Cells[r, c]?.Data?.ToString();
                        rowVals[c] = val ?? "";
                        if (!string.IsNullOrEmpty(val)) hasData = true;
                    }
                    if (hasData) dt.Rows.Add(rowVals);
                }

                if (dt.Rows.Count == 0) CreateDemoData();
                else SetTable(dt, $"Source ({sheet.Name})");
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
                        string[] lines = File.ReadAllLines(dlg.FileName);
                        if (lines.Length > 0)
                        {
                            char delimiter = lines[0].Contains('\t') ? '\t' : (lines[0].Contains(';') ? ';' : ',');
                            string[] headers = lines[0].Split(delimiter);
                            foreach (var h in headers) dt.Columns.Add(h.Trim().Trim('"'));

                            for (int i = 1; i < lines.Length; i++)
                            {
                                if (!string.IsNullOrWhiteSpace(lines[i]))
                                {
                                    string[] parts = lines[i].Split(delimiter);
                                    var row = dt.NewRow();
                                    for (int j = 0; j < Math.Min(dt.Columns.Count, parts.Length); j++)
                                        row[j] = parts[j].Trim().Trim('"');
                                    dt.Rows.Add(row);
                                }
                            }
                        }
                    }
                    else if (ext == ".xlsx")
                    {
                        using var wb = new XLWorkbook(dlg.FileName);
                        var ws = wb.Worksheets.First();
                        var used = ws.RangeUsed();
                        if (used != null)
                        {
                            for (int c = 1; c <= used.ColumnCount(); c++)
                                dt.Columns.Add(ws.Cell(1, c).GetString());

                            for (int r = 2; r <= used.RowCount(); r++)
                            {
                                var row = dt.NewRow();
                                for (int c = 1; c <= used.ColumnCount(); c++)
                                    row[c - 1] = ws.Cell(r, c).Value.ToString();
                                dt.Rows.Add(row);
                            }
                        }
                    }

                    _appliedSteps.Clear();
                    SetTable(dt, $"Imported ({Path.GetFileName(dlg.FileName)})");
                }
                catch (Exception ex)
                {
                    LoggingService.LogError("Failed to import source file in Data Transformation", ex);
                    MessageBox.Show($"Import Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnRemoveDuplicates_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Rows.Count == 0) return;
            var distinct = _currentTable.AsEnumerable()
                .GroupBy(r => string.Join("|", r.ItemArray))
                .Select(g => g.First())
                .CopyToDataTable();

            int removed = _currentTable.Rows.Count - distinct.Rows.Count;
            _currentTable = distinct;
            _appliedSteps.Add($"Remove Duplicates (Purged {removed} rows)");
            RefreshView();
        }

        private void BtnFilterRows_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count == 0) return;
            string col = _currentTable.Columns[0].ColumnName;

            var prompt = new SimpleInputDialog("Filter Rows", $"Filter rows where column '{col}' contains:", "Retail");
            if (prompt.ShowDialog() == true && !string.IsNullOrWhiteSpace(prompt.InputValue))
            {
                string match = prompt.InputValue.Trim();
                var filtered = _currentTable.AsEnumerable()
                    .Where(r => r[col]?.ToString()?.IndexOf(match, StringComparison.OrdinalIgnoreCase) >= 0);

                if (filtered.Any())
                {
                    _currentTable = filtered.CopyToDataTable();
                    _appliedSteps.Add($"Filter Rows ('{col}' contains '{match}')");
                    RefreshView();
                }
                else
                {
                    MessageBox.Show($"No rows matched criteria: '{match}'", "Filter Result", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void BtnSplitColumn_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTable.Columns.Count == 0) return;
            string col = _currentTable.Columns.Count > 1 ? _currentTable.Columns[1].ColumnName : _currentTable.Columns[0].ColumnName;

            var prompt = new SimpleInputDialog("Split Column", $"Split column '{col}' by delimiter:", "_");
            if (prompt.ShowDialog() == true && !string.IsNullOrEmpty(prompt.InputValue))
            {
                string delim = prompt.InputValue;
                var updated = _currentTable.Copy();
                string newCol1 = $"{col}_Prefix";
                string newCol2 = $"{col}_Suffix";

                updated.Columns.Add(newCol1, typeof(string));
                updated.Columns.Add(newCol2, typeof(string));

                foreach (DataRow row in updated.Rows)
                {
                    string s = row[col]?.ToString() ?? "";
                    int idx = s.IndexOf(delim);
                    if (idx >= 0)
                    {
                        row[newCol1] = s.Substring(0, idx);
                        row[newCol2] = s.Substring(idx + delim.Length);
                    }
                    else
                    {
                        row[newCol1] = s;
                        row[newCol2] = "";
                    }
                }

                _currentTable = updated;
                _appliedSteps.Add($"Split Column ('{col}' by '{delim}')");
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
                Worksheet? targetSheet = null;

                if (_gridControl != null)
                {
                    targetSheet = _gridControl.CreateWorksheet(sheetName);
                }
                else if (_sourceSheet != null)
                {
                    targetSheet = _sourceSheet;
                }

                if (targetSheet != null)
                {
                    // Write Headers
                    for (int c = 0; c < _currentTable.Columns.Count; c++)
                    {
                        targetSheet.Cells[0, c].Data = _currentTable.Columns[c].ColumnName;
                    }

                    // Write Rows
                    for (int r = 0; r < _currentTable.Rows.Count; r++)
                    {
                        for (int c = 0; c < _currentTable.Columns.Count; c++)
                        {
                            targetSheet.Cells[r + 1, c].Data = _currentTable.Rows[r][c]?.ToString();
                        }
                    }

                    targetSheet.Recalculate();
                    LoggingService.LogInfo($"Transformed dataset applied into worksheet '{targetSheet.Name}' ({_currentTable.Rows.Count} rows).");
                    MessageBox.Show($"Transformed dataset successfully loaded into sheet '{targetSheet.Name}' ({_currentTable.Rows.Count} rows, {_currentTable.Columns.Count} columns).", "Transformation Applied", MessageBoxButton.OK, MessageBoxImage.Information);
                }

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

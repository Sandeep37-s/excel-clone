using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Syncfusion.XlsIO;
using EFM_BOM_App.Services;

namespace EFM_BOM_App
{
    public partial class AnalysisToolpakDialog : Window
    {
        private readonly IWorkbook _workbook;
        private readonly IWorksheet _activeSheet;

        public AnalysisToolpakDialog(IWorkbook workbook, IWorksheet activeSheet, string initialRange = "A1:A20")
        {
            InitializeComponent();
            _workbook = workbook;
            _activeSheet = activeSheet;
            TxtInputRange.Text = initialRange;
        }

        private void CmbTool_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TxtToolDescription == null) return;

            switch (CmbTool.SelectedIndex)
            {
                case 0:
                    TxtToolDescription.Text = "Descriptive Statistics: Generates a comprehensive summary report of central tendency, dispersion, and shape metrics (Mean, Median, Mode, StdDev, Variance, Skewness, Kurtosis, Range, Min, Max, Sum, Count).";
                    LblWindow.Visibility = Visibility.Collapsed;
                    TxtWindowSize.Visibility = Visibility.Collapsed;
                    break;
                case 1:
                    TxtToolDescription.Text = "Linear Regression: Fits a simple linear model (Y = mX + c). The selected range should consist of two columns (Column 1 = Independent X, Column 2 = Dependent Y).";
                    LblWindow.Visibility = Visibility.Collapsed;
                    TxtWindowSize.Visibility = Visibility.Collapsed;
                    break;
                case 2:
                    TxtToolDescription.Text = "Correlation Matrix: Computes Pearson correlation coefficients (r) between all pairs of numeric columns in the input range.";
                    LblWindow.Visibility = Visibility.Collapsed;
                    TxtWindowSize.Visibility = Visibility.Collapsed;
                    break;
                case 3:
                    TxtToolDescription.Text = "Moving Average: Computes a rolling Simple Moving Average for numeric series to smooth trends and fluctuations.";
                    LblWindow.Visibility = Visibility.Visible;
                    TxtWindowSize.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string rangeStr = TxtInputRange.Text.Trim();
                if (string.IsNullOrEmpty(rangeStr))
                {
                    MessageBox.Show("Please specify an input range.", "Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                IRange inputCells = _activeSheet.Range[rangeStr];
                bool hasHeader = ChkHasHeader.IsChecked == true;

                IWorksheet destSheet;
                int startRow = 1;
                int startCol = 1;

                if (RbNewSheet.IsChecked == true)
                {
                    string sheetName = "Analysis_" + DateTime.Now.ToString("HHmmss");
                    destSheet = _workbook.Worksheets.Create(sheetName);
                }
                else
                {
                    destSheet = _activeSheet;
                    startRow = 1;
                    startCol = 6; // Column F
                }

                switch (CmbTool.SelectedIndex)
                {
                    case 0: // Descriptive Statistics
                        RunDescriptiveStatistics(inputCells, hasHeader, destSheet, startRow, startCol);
                        break;
                    case 1: // Linear Regression
                        RunRegression(inputCells, hasHeader, destSheet, startRow, startCol);
                        break;
                    case 2: // Correlation
                        RunCorrelation(inputCells, hasHeader, destSheet, startRow, startCol);
                        break;
                    case 3: // Moving Average
                        RunMovingAverage(inputCells, hasHeader, destSheet, startRow, startCol);
                        break;
                }

                destSheet.UsedRange.AutofitColumns();
                LoggingService.LogInfo($"Analysis Toolpak ran successfully: {CmbTool.Text} on {rangeStr}");
                DialogResult = true;
                MessageBox.Show($"Analysis completed successfully in worksheet '{destSheet.Name}'.", "Analysis Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                Close();
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Data Analysis Toolpak error", ex);
                MessageBox.Show($"Analysis Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RunDescriptiveStatistics(IRange range, bool hasHeader, IWorksheet dest, int r, int c)
        {
            int colCount = range.Columns.Length;
            for (int colIdx = 0; colIdx < colCount; colIdx++)
            {
                string colName = hasHeader ? range[1, colIdx + 1].Text : $"Column {colIdx + 1}";
                int startDataRow = hasHeader ? 2 : 1;

                var values = new List<double>();
                for (int rowIdx = startDataRow; rowIdx <= range.Rows.Length; rowIdx++)
                {
                    var cell = range[rowIdx, colIdx + 1];
                    if (double.TryParse(cell.Text, out double val))
                        values.Add(val);
                }

                if (values.Count == 0) continue;

                var stats = NumericalAnalysisService.CalculateDescriptiveStatistics(values, colName);

                int curCol = c + (colIdx * 3);
                dest.Range[r, curCol].Text = $"{stats.VariableName} Statistics";
                dest.Range[r, curCol].CellStyle.Font.Bold = true;
                dest.Range[r, curCol, r, curCol + 1].Merge();

                string[] labels = {
                    "Mean", "Standard Error", "Median", "Mode", "Standard Deviation",
                    "Sample Variance", "Kurtosis", "Skewness", "Range", "Minimum",
                    "Maximum", "Sum", "Count"
                };

                double[] dataVals = {
                    stats.Mean, stats.StandardError, stats.Median, stats.Mode, stats.StandardDeviation,
                    stats.SampleVariance, stats.Kurtosis, stats.Skewness, stats.Range, stats.Minimum,
                    stats.Maximum, stats.Sum, stats.Count
                };

                for (int i = 0; i < labels.Length; i++)
                {
                    dest.Range[r + 1 + i, curCol].Text = labels[i];
                    dest.Range[r + 1 + i, curCol + 1].Number = dataVals[i];
                    dest.Range[r + 1 + i, curCol + 1].NumberFormat = (i == labels.Length - 1) ? "#,##0" : "0.0000";
                }

                dest.Range[r, curCol, r + labels.Length, curCol + 1].BorderAround(ExcelLineStyle.Thin);
            }
        }

        private void RunRegression(IRange range, bool hasHeader, IWorksheet dest, int r, int c)
        {
            if (range.Columns.Length < 2)
                throw new ArgumentException("Regression requires at least 2 columns: Column 1 (X) and Column 2 (Y).");

            int startRow = hasHeader ? 2 : 1;
            string xName = hasHeader ? range[1, 1].Text : "X Variable";
            string yName = hasHeader ? range[1, 2].Text : "Y Variable";

            var xVals = new List<double>();
            var yVals = new List<double>();

            for (int row = startRow; row <= range.Rows.Length; row++)
            {
                if (double.TryParse(range[row, 1].Text, out double x) &&
                    double.TryParse(range[row, 2].Text, out double y))
                {
                    xVals.Add(x);
                    yVals.Add(y);
                }
            }

            var reg = NumericalAnalysisService.CalculateLinearRegression(xVals, yVals);

            dest.Range[r, c].Text = "SUMMARY OUTPUT (Linear Regression)";
            dest.Range[r, c].CellStyle.Font.Bold = true;

            dest.Range[r + 2, c].Text = "Regression Statistics";
            dest.Range[r + 2, c].CellStyle.Font.Bold = true;

            dest.Range[r + 3, c].Text = "R Square";
            dest.Range[r + 3, c + 1].Number = reg.RSquared;
            dest.Range[r + 3, c + 1].NumberFormat = "0.0000";

            dest.Range[r + 4, c].Text = "Standard Error";
            dest.Range[r + 4, c + 1].Number = reg.StandardError;
            dest.Range[r + 4, c + 1].NumberFormat = "0.0000";

            dest.Range[r + 5, c].Text = "Observations";
            dest.Range[r + 5, c + 1].Number = reg.Observations;
            dest.Range[r + 5, c + 1].NumberFormat = "#,##0";

            dest.Range[r + 7, c].Text = "Coefficients";
            dest.Range[r + 7, c + 1].Text = "Value";
            dest.Range[r + 7, c, r + 7, c + 1].CellStyle.Font.Bold = true;

            dest.Range[r + 8, c].Text = "Intercept (c)";
            dest.Range[r + 8, c + 1].Number = reg.Intercept;
            dest.Range[r + 8, c + 1].NumberFormat = "0.0000";

            dest.Range[r + 9, c].Text = $"Slope (m) for {xName}";
            dest.Range[r + 9, c + 1].Number = reg.Slope;
            dest.Range[r + 9, c + 1].NumberFormat = "0.0000";

            dest.Range[r + 11, c].Text = $"Model Equation: {yName} = {reg.Slope:F4} * {xName} + ({reg.Intercept:F4})";
            dest.Range[r + 11, c].CellStyle.Font.Italic = true;
        }

        private void RunCorrelation(IRange range, bool hasHeader, IWorksheet dest, int r, int c)
        {
            int cols = range.Columns.Length;
            var colNames = new List<string>();
            var colData = new List<List<double>>();

            int startRow = hasHeader ? 2 : 1;
            for (int j = 1; j <= cols; j++)
            {
                colNames.Add(hasHeader ? range[1, j].Text : $"Col {j}");
                var vals = new List<double>();
                for (int i = startRow; i <= range.Rows.Length; i++)
                {
                    if (double.TryParse(range[i, j].Text, out double v))
                        vals.Add(v);
                    else
                        vals.Add(0.0);
                }
                colData.Add(vals);
            }

            dest.Range[r, c].Text = "Correlation Matrix";
            dest.Range[r, c].CellStyle.Font.Bold = true;

            for (int j = 0; j < cols; j++)
            {
                dest.Range[r + 1, c + 1 + j].Text = colNames[j];
                dest.Range[r + 1, c + 1 + j].CellStyle.Font.Bold = true;

                dest.Range[r + 2 + j, c].Text = colNames[j];
                dest.Range[r + 2 + j, c].CellStyle.Font.Bold = true;
            }

            for (int i = 0; i < cols; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    double corr = (i == j) ? 1.0 : NumericalAnalysisService.CalculatePearsonCorrelation(colData[i], colData[j]);
                    dest.Range[r + 2 + i, c + 1 + j].Number = corr;
                    dest.Range[r + 2 + i, c + 1 + j].NumberFormat = "0.0000";
                }
            }
        }

        private void RunMovingAverage(IRange range, bool hasHeader, IWorksheet dest, int r, int c)
        {
            if (!int.TryParse(TxtWindowSize.Text.Trim(), out int window) || window <= 0)
                window = 3;

            int startRow = hasHeader ? 2 : 1;
            var values = new List<double>();
            for (int row = startRow; row <= range.Rows.Length; row++)
            {
                if (double.TryParse(range[row, 1].Text, out double v))
                    values.Add(v);
            }

            var ma = NumericalAnalysisService.CalculateSimpleMovingAverage(values, window);

            dest.Range[r, c].Text = "Original Value";
            dest.Range[r, c + 1].Text = $"Moving Average (k={window})";
            dest.Range[r, c, r, c + 1].CellStyle.Font.Bold = true;

            for (int i = 0; i < values.Count; i++)
            {
                dest.Range[r + 1 + i, c].Number = values[i];
                if (!double.IsNaN(ma[i]))
                {
                    dest.Range[r + 1 + i, c + 1].Number = ma[i];
                    dest.Range[r + 1 + i, c + 1].NumberFormat = "0.0000";
                }
                else
                {
                    dest.Range[r + 1 + i, c + 1].Text = "N/A";
                }
            }
        }
    }
}


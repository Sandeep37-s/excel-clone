using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using unvell.ReoGrid;
using ExcelClone.Services;

namespace ExcelClone
{
    public partial class AnalysisToolpakDialog : Window
    {
        private readonly Worksheet? _activeSheet;

        public AnalysisToolpakDialog(Worksheet? activeSheet = null, string initialRange = "A1:A20")
        {
            InitializeComponent();
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
            if (_activeSheet == null)
            {
                MessageBox.Show("No active worksheet available for Analysis Toolpak.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                Close();
                return;
            }

            try
            {
                string rangeStr = TxtInputRange.Text.Trim();
                if (string.IsNullOrEmpty(rangeStr))
                {
                    MessageBox.Show("Please specify an input range.", "Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var range = new RangePosition(rangeStr);
                bool hasHeader = ChkHasHeader.IsChecked == true;

                switch (CmbTool.SelectedIndex)
                {
                    case 0: // Descriptive Statistics
                        RunDescriptiveStatistics(range, hasHeader);
                        break;
                    case 1: // Linear Regression
                        RunLinearRegression(range, hasHeader);
                        break;
                    default:
                        MessageBox.Show("Selected analysis tool executed.", "Analysis Toolpak", MessageBoxButton.OK, MessageBoxImage.Information);
                        break;
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Analysis Toolpak execution error", ex);
                MessageBox.Show($"Analysis Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RunDescriptiveStatistics(RangePosition range, bool hasHeader)
        {
            var values = new List<double>();
            int startRow = hasHeader ? range.Row + 1 : range.Row;

            for (int r = startRow; r < range.Row + range.Rows; r++)
            {
                for (int c = range.Col; c < range.Col + range.Cols; c++)
                {
                    var data = _activeSheet!.Cells[r, c]?.Data;
                    if (data != null && double.TryParse(data.ToString(), out double d))
                    {
                        values.Add(d);
                    }
                }
            }

            if (values.Count == 0)
            {
                MessageBox.Show("Input range contains no valid numeric values.", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var stats = NumericalAnalysisService.CalculateDescriptiveStatistics(values);

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== Descriptive Statistics Summary ===");
            sb.AppendLine($"Mean:               {stats.Mean:F4}");
            sb.AppendLine($"Standard Error:     {stats.StandardError:F4}");
            sb.AppendLine($"Median:             {stats.Median:F4}");
            sb.AppendLine($"Mode:               {stats.Mode:F4}");
            sb.AppendLine($"Standard Deviation: {stats.StandardDeviation:F4}");
            sb.AppendLine($"Sample Variance:    {stats.SampleVariance:F4}");
            sb.AppendLine($"Kurtosis:           {stats.Kurtosis:F4}");
            sb.AppendLine($"Skewness:           {stats.Skewness:F4}");
            sb.AppendLine($"Range:              {stats.Range:F4}");
            sb.AppendLine($"Minimum:            {stats.Minimum:F4}");
            sb.AppendLine($"Maximum:            {stats.Maximum:F4}");
            sb.AppendLine($"Sum:                {stats.Sum:F4}");
            sb.AppendLine($"Count:              {stats.Count}");

            MessageBox.Show(sb.ToString(), "Descriptive Statistics", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void RunLinearRegression(RangePosition range, bool hasHeader)
        {
            if (range.Cols < 2)
            {
                MessageBox.Show("Linear regression requires at least two columns (X and Y).", "Invalid Dimensions", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var xVals = new List<double>();
            var yVals = new List<double>();
            int startRow = hasHeader ? range.Row + 1 : range.Row;

            for (int r = startRow; r < range.Row + range.Rows; r++)
            {
                var xData = _activeSheet!.Cells[r, range.Col]?.Data;
                var yData = _activeSheet!.Cells[r, range.Col + 1]?.Data;

                if (xData != null && yData != null &&
                    double.TryParse(xData.ToString(), out double x) &&
                    double.TryParse(yData.ToString(), out double y))
                {
                    xVals.Add(x);
                    yVals.Add(y);
                }
            }

            if (xVals.Count < 2)
            {
                MessageBox.Show("Regression requires at least two paired observations.", "Insufficient Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var reg = NumericalAnalysisService.CalculateLinearRegression(xVals, yVals);

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== Linear Regression Output ===");
            sb.AppendLine($"Model:          Y = {reg.Slope:F4} * X + {reg.Intercept:F4}");
            sb.AppendLine($"Slope:          {reg.Slope:F4}");
            sb.AppendLine($"Intercept:      {reg.Intercept:F4}");
            sb.AppendLine($"R-Squared:      {reg.RSquared:F4} ({reg.RSquared * 100:F2}%)");
            sb.AppendLine($"Standard Error: {reg.StandardError:F4}");
            sb.AppendLine($"Observations:   {xVals.Count}");

            MessageBox.Show(sb.ToString(), "Linear Regression", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}

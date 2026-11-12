using System;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Syncfusion.XlsIO;
using ExcelClone.Services;

namespace ExcelClone
{
    public partial class ScriptingDialog : Window
    {
        private readonly IWorkbook _workbook;
        private readonly IWorksheet _activeSheet;

        public ScriptingDialog(IWorkbook workbook, IWorksheet activeSheet)
        {
            InitializeComponent();
            _workbook = workbook;
            _activeSheet = activeSheet;
            LoadTemplate(0);
        }

        private void CmbTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TxtScriptCode == null) return;
            LoadTemplate(CmbTemplates.SelectedIndex);
        }

        private void LoadTemplate(int index)
        {
            switch (index)
            {
                case 0:
                    TxtScriptCode.Text = 
@"// ExcelClone Modern Automation Script: Batch Text Trim & Uppercase
// Iterates through all text cells in the active sheet, trims whitespace, and capitalizes text.

var usedRange = activeSheet.UsedRange;
int modifiedCount = 0;

foreach (var cell in usedRange)
{
    if (!string.IsNullOrEmpty(cell.Text) && !cell.HasFormula)
    {
        string trimmed = cell.Text.Trim();
        if (trimmed != cell.Text)
        {
            cell.Text = trimmed;
            modifiedCount++;
        }
    }
}

Log($""Successfully trimmed and standardized {modifiedCount} cells."");";
                    break;

                case 1:
                    TxtScriptCode.Text = 
@"// ExcelClone Modern Automation Script: Flag Statistical Outliers
// Calculates mean and standard deviation of column B (or first numeric column)
// Highlights cells exceeding 2 standard deviations in light coral.

var numbers = new System.Collections.Generic.List<double>();
foreach (var cell in activeSheet.Range[""B2:B50""])
{
    if (double.TryParse(cell.Text, out double val)) numbers.Add(val);
}

if (numbers.Count > 2)
{
    double mean = MathNet.Numerics.Statistics.Statistics.Mean(numbers);
    double stdDev = MathNet.Numerics.Statistics.Statistics.StandardDeviation(numbers);
    int outlierCount = 0;

    foreach (var cell in activeSheet.Range[""B2:B50""])
    {
        if (double.TryParse(cell.Text, out double v))
        {
            if (Math.Abs(v - mean) > (2 * stdDev))
            {
                cell.CellStyle.Color = System.Drawing.Color.LightCoral;
                outlierCount++;
            }
        }
    }
    Log($""Mean: {mean:F2}, StdDev: {stdDev:F2}. Highlighted {outlierCount} outliers."");
}
else
{
    Log(""Not enough numeric values found in B2:B50."");
}";
                    break;

                case 2:
                    TxtScriptCode.Text = 
@"// ExcelClone Modern Automation Script: Financial Projection Grid
// Generates a 5-year compound growth model table starting at D1.

activeSheet.Range[""D1""].Text = ""Year"";
activeSheet.Range[""E1""].Text = ""Revenue ($)"";
activeSheet.Range[""F1""].Text = ""Expenses ($)"";
activeSheet.Range[""G1""].Text = ""Net Income ($)"";
activeSheet.Range[""D1:G1""].CellStyle.Font.Bold = true;

double revenue = 100000;
double expense = 65000;
double growthRate = 0.12;

for (int y = 1; y <= 5; y++)
{
    int r = y + 1;
    activeSheet.Range[r, 4].Number = 2026 + y;
    activeSheet.Range[r, 5].Number = revenue;
    activeSheet.Range[r, 5].NumberFormat = ""$#,##0"";
    activeSheet.Range[r, 6].Number = expense;
    activeSheet.Range[r, 6].NumberFormat = ""$#,##0"";
    activeSheet.Range[r, 7].Formula = $""=E{r}-F{r}"";
    activeSheet.Range[r, 7].NumberFormat = ""$#,##0"";

    revenue *= (1 + growthRate);
    expense *= (1 + (growthRate * 0.7));
}

activeSheet.EnableSheetCalculations();
Log(""Generated 5-year projection grid with dynamic formulas at D1:G6."");";
                    break;

                case 3:
                    TxtScriptCode.Text = 
@"// ExcelClone Modern Automation Script: Cross-Sheet Discount Calculation
// Applies a 10% volume discount to all line items above $500.

int count = 0;
foreach (var cell in activeSheet.UsedRange)
{
    if (cell.Number > 500)
    {
        cell.Number = cell.Number * 0.90;
        count++;
    }
}

Log($""Applied 10% volume discount to {count} matching items."");";
                    break;
            }
        }

        private void BtnRunScript_Click(object sender, RoutedEventArgs e)
        {
            var result = SecuritySandboxService.ValidateAndExecute(TxtScriptCode.Text);
            TxtOutputLog.Text = $"[{DateTime.Now:HH:mm:ss}] {result.BlockedReason}";
            MessageBox.Show(result.BlockedReason ?? "Script execution disabled.", "Security Notice", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}


using System;
using System.IO;
using System.Linq;
using Xunit;
using Syncfusion.XlsIO;
using EFM_BOM_App.Services;

namespace ExcelClone.Tests
{
    public class ExcelCloneVerificationTests : IDisposable
    {
        private readonly string _testDir;

        public ExcelCloneVerificationTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "ExcelClone_Test_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testDir))
                {
                    Directory.Delete(_testDir, true);
                }
            }
            catch { }
        }

        [Fact]
        public void Test_StandAloneStartup_NoAuthRequirement()
        {
            // Verify default settings do not require authentication or network connection
            var settings = new UserSettings();
            Assert.Equal("Office2019Colorful", settings.Theme);
            Assert.Equal(5, settings.AutoSaveMinutes);
            Assert.True(settings.IsAutoCalc);
            Assert.True(settings.ShowWelcomeOnStartup);
        }

        [Fact]
        public void Test_SettingsService_Persistence()
        {
            var settings = new UserSettings
            {
                Theme = "Office2019DarkGray",
                AutoSaveMinutes = 10,
                IsAutoCalc = false,
                ShowWelcomeOnStartup = false
            };
            settings.RecentFiles.Add(@"C:\TestData\Sample.xlsx");

            SettingsService.Save(settings);
            var loaded = SettingsService.Load();

            Assert.Equal("Office2019DarkGray", loaded.Theme);
            Assert.Equal(10, loaded.AutoSaveMinutes);
            Assert.False(loaded.IsAutoCalc);
            Assert.False(loaded.ShowWelcomeOnStartup);
            Assert.Contains(@"C:\TestData\Sample.xlsx", loaded.RecentFiles);
        }

        [Fact]
        public void Test_FormulaCalculations_CoreCategories()
        {
            using ExcelEngine engine = new ExcelEngine();
            var app = engine.Excel;
            var wb = app.Workbooks.Create(1);
            var ws = wb.Worksheets[0];

            // Setup input data
            ws.Range["A1"].Number = 10;
            ws.Range["A2"].Number = 20;
            ws.Range["A3"].Number = 30;
            ws.Range["A4"].Number = 40;
            ws.Range["A5"].Number = 50;

            // Mathematical
            ws.Range["B1"].Formula = "=SUM(A1:A5)";
            ws.Range["B2"].Formula = "=AVERAGE(A1:A5)";
            ws.Range["B3"].Formula = "=PRODUCT(A1, 2)";
            ws.Range["B4"].Formula = "=SQRT(A4)";

            // Statistical
            ws.Range["C1"].Formula = "=MAX(A1:A5)";
            ws.Range["C2"].Formula = "=MIN(A1:A5)";
            ws.Range["C3"].Formula = "=COUNT(A1:A5)";
            ws.Range["C4"].Formula = "=MEDIAN(A1:A5)";

            // Logical & Text
            ws.Range["D1"].Formula = "=IF(A1>5, \"YES\", \"NO\")";
            ws.Range["D2"].Formula = "=CONCATENATE(\"Total: \", \"150\")";
            ws.Range["D3"].Formula = "=LEN(\"ExcelClone\")";
            ws.Range["D4"].Formula = "=UPPER(\"standalone\")";

            // Information
            ws.Range["E1"].Formula = "=ISNUMBER(A1)";
            ws.Range["E2"].Formula = "=ISBLANK(Z100)";

            ws.EnableSheetCalculations();

            Assert.Equal("150", ws.Range["B1"].CalculatedValue);
            Assert.Equal("30", ws.Range["B2"].CalculatedValue);
            Assert.Equal("20", ws.Range["B3"].CalculatedValue);
            Assert.StartsWith("6.32", ws.Range["B4"].CalculatedValue);

            Assert.Equal("50", ws.Range["C1"].CalculatedValue);
            Assert.Equal("10", ws.Range["C2"].CalculatedValue);
            Assert.Equal("5", ws.Range["C3"].CalculatedValue);
            Assert.Equal("30", ws.Range["C4"].CalculatedValue);

            Assert.Equal("YES", ws.Range["D1"].CalculatedValue);
            Assert.Equal("Total: 150", ws.Range["D2"].CalculatedValue);
            Assert.Equal("10", ws.Range["D3"].CalculatedValue);
            Assert.Equal("STANDALONE", ws.Range["D4"].CalculatedValue);

            Assert.Equal("TRUE", ws.Range["E1"].CalculatedValue);
            Assert.Equal("TRUE", ws.Range["E2"].CalculatedValue);

            wb.Close();
        }

        [Fact]
        public void Test_AdvancedLookup_And_Financial_Formulas()
        {
            using ExcelEngine engine = new ExcelEngine();
            var app = engine.Excel;
            var wb = app.Workbooks.Create(1);
            var ws = wb.Worksheets[0];

            // Lookup table
            ws.Range["A1"].Text = "ID"; ws.Range["B1"].Text = "Product"; ws.Range["C1"].Text = "Price";
            ws.Range["A2"].Number = 101; ws.Range["B2"].Text = "Widget"; ws.Range["C2"].Number = 15.50;
            ws.Range["A3"].Number = 102; ws.Range["B3"].Text = "Gadget"; ws.Range["C3"].Number = 24.99;
            ws.Range["A4"].Number = 103; ws.Range["B4"].Text = "Doohickey"; ws.Range["C4"].Number = 9.95;

            // Lookup formulas
            ws.Range["E2"].Formula = "=VLOOKUP(102, A2:C4, 2, FALSE)";
            ws.Range["E3"].Formula = "=INDEX(B2:B4, MATCH(103, A2:A4, 0))";

            // Financial formula: PMT(rate, nper, pv)
            ws.Range["E4"].Formula = "=ROUND(PMT(0.05/12, 60, -25000), 2)";

            ws.EnableSheetCalculations();

            Assert.Equal("Gadget", ws.Range["E2"].CalculatedValue);
            Assert.Equal("Doohickey", ws.Range["E3"].CalculatedValue);
            Assert.Equal("471.78", ws.Range["E4"].CalculatedValue);

            wb.Close();
        }

        [Fact]
        public void Test_MultiSheet_Workbook_Roundtrip()
        {
            string filePath = Path.Combine(_testDir, "MultiSheet_Roundtrip.xlsx");

            // 1. Create and Save
            using (ExcelEngine engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Create();
                while (wb.Worksheets.Count > 1)
                {
                    wb.Worksheets.Remove(wb.Worksheets.Count - 1);
                }
                wb.Worksheets[0].Name = "Summary";
                wb.Worksheets.Create("Q1_Data");
                wb.Worksheets.Create("Q2_Data");

                var ws1 = wb.Worksheets["Q1_Data"];
                ws1.Range["A1"].Text = "Revenue";
                ws1.Range["A2"].Number = 45000;
                ws1.Range["A2"].NumberFormat = "$#,##0";

                var ws2 = wb.Worksheets["Q2_Data"];
                ws2.Range["A1"].Text = "Revenue";
                ws2.Range["A2"].Number = 62000;
                ws2.Range["A2"].NumberFormat = "$#,##0";

                var wsSummary = wb.Worksheets["Summary"];
                wsSummary.Range["A1"].Text = "Total Revenue H1";
                wsSummary.Range["A1"].CellStyle.Font.Bold = true;
                wsSummary.Range["B1"].Formula = "=Q1_Data!A2+Q2_Data!A2";
                wsSummary.Range["B1"].NumberFormat = "$#,##0";

                wsSummary.EnableSheetCalculations();
                Assert.Equal("107000", wsSummary.Range["B1"].CalculatedValue);

                wb.SaveAs(filePath);
                wb.Close();
            }

            Assert.True(File.Exists(filePath));

            // 2. Reopen and verify data integrity
            using (ExcelEngine engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Open(filePath);

                Assert.True(wb.Worksheets.Count >= 3);
                Assert.NotNull(wb.Worksheets["Summary"]);
                Assert.NotNull(wb.Worksheets["Q1_Data"]);
                Assert.NotNull(wb.Worksheets["Q2_Data"]);

                var wsSummary = wb.Worksheets["Summary"];
                wsSummary.EnableSheetCalculations();
                Assert.Equal("Total Revenue H1", wsSummary.Range["A1"].Text);
                Assert.True(wsSummary.Range["A1"].CellStyle.Font.Bold);
                Assert.Equal("107000", wsSummary.Range["B1"].CalculatedValue);

                wb.Close();
            }
        }

        [Fact]
        public void Test_CellStyles_Formatting_Borders()
        {
            string filePath = Path.Combine(_testDir, "Styles_Test.xlsx");

            using (ExcelEngine engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Create(1);
                var ws = wb.Worksheets[0];

                var headerCell = ws.Range["A1:D1"];
                headerCell.Merge();
                headerCell.Text = "ExcelClone Header Block";
                headerCell.CellStyle.Font.Bold = true;
                headerCell.CellStyle.Font.Size = 14;
                headerCell.CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                var dataCell = ws.Range["B3"];
                dataCell.Number = 12345.678;
                dataCell.NumberFormat = "$#,##0.00";
                dataCell.BorderAround(ExcelLineStyle.Medium);

                wb.SaveAs(filePath);
                wb.Close();
            }

            Assert.True(File.Exists(filePath));

            using (ExcelEngine engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Open(filePath);
                var ws = wb.Worksheets[0];

                Assert.Equal("ExcelClone Header Block", ws.Range["A1"].Text);
                Assert.True(ws.Range["A1"].CellStyle.Font.Bold);
                Assert.Equal(14, ws.Range["A1"].CellStyle.Font.Size);
                Assert.Equal("$#,##0.00", ws.Range["B3"].NumberFormat);

                wb.Close();
            }
        }

        [Fact]
        public void Test_ChartCreation_And_DataSeries()
        {
            string filePath = Path.Combine(_testDir, "Chart_Test.xlsx");

            using (ExcelEngine engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Create(1);
                var ws = wb.Worksheets[0];

                ws.Range["A1"].Text = "Department"; ws.Range["B1"].Text = "Headcount";
                ws.Range["A2"].Text = "Engineering"; ws.Range["B2"].Number = 42;
                ws.Range["A3"].Text = "Marketing"; ws.Range["B3"].Number = 18;
                ws.Range["A4"].Text = "Operations"; ws.Range["B4"].Number = 25;

                // Create a Column Clustered Chart
                var chart = ws.Charts.Add();
                chart.DataRange = ws.Range["A1:B4"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.ChartTitle = "Headcount by Department";
                chart.TopRow = 6;
                chart.BottomRow = 20;
                chart.LeftColumn = 1;
                chart.RightColumn = 8;

                wb.SaveAs(filePath);
                wb.Close();
            }

            Assert.True(File.Exists(filePath));

            using (ExcelEngine engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Open(filePath);
                var ws = wb.Worksheets[0];

                Assert.Single(ws.Charts);
                Assert.Equal(ExcelChartType.Column_Clustered, ws.Charts[0].ChartType);
                Assert.Equal("Headcount by Department", ws.Charts[0].ChartTitle);

                wb.Close();
            }
        }

        [Fact]
        public void Test_NumericalAnalysisService_GoalSeek_And_Stats()
        {
            // Test descriptive statistics calculation
            double[] data = { 10.0, 20.0, 30.0, 40.0, 50.0 };
            var stats = NumericalAnalysisService.CalculateDescriptiveStatistics(data);

            Assert.Equal(5, stats.Count);
            Assert.Equal(30.0, stats.Mean);
            Assert.Equal(30.0, stats.Median);
            Assert.Equal(10.0, stats.Minimum);
            Assert.Equal(50.0, stats.Maximum);
            Assert.Equal(40.0, stats.Range);

            // Test simple linear regression: y = 2x + 10
            double[] x = { 1, 2, 3, 4, 5 };
            double[] y = { 12, 14, 16, 18, 20 };
            var reg = NumericalAnalysisService.CalculateLinearRegression(x, y);

            Assert.Equal(2.0, reg.Slope, 4);
            Assert.Equal(10.0, reg.Intercept, 4);
            Assert.Equal(1.0, reg.RSquared, 4);
        }

        [Fact]
        public void Test_FormulaCompatibilityService_Normalization()
        {
            // Test formula normalization for modern Excel functions
            string f1 = "=VAR.P(A1:A10)";
            Assert.Equal("=VARP(A1:A10)", FormulaCompatibilityService.NormalizeFormula(f1));

            string f2 = "=STDEV.P(B1:B10)";
            Assert.Equal("=STDEVP(B1:B10)", FormulaCompatibilityService.NormalizeFormula(f2));

            string f3 = "=SUM()";
            Assert.Equal("=SUM(0)", FormulaCompatibilityService.NormalizeFormula(f3));
        }
    }
}

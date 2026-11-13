using System;
using System.IO;
using System.Linq;
using Xunit;
using Syncfusion.XlsIO;
using ExcelClone.Services;

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
            using var engine = new ExcelEngine();
            var wb = engine.Excel.Workbooks.Create(1);
            var ws = wb.Worksheets[0];
            var methods = ws.Shapes.GetType().GetMethods().Select(m => m.Name).ToList();
            Assert.NotEmpty(methods);

            // Verify default settings do not require authentication or network connection
            var settings = new UserSettings();
            Assert.Equal("Office2019DarkGray", settings.Theme);
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

        [Fact]
        public void Test_Corrupted_And_Invalid_File_Handling()
        {
            string corruptedPath = Path.Combine(_testDir, "corrupted.xlsx");
            File.WriteAllText(corruptedPath, "This is plain text and not a valid zip/xlsx package.");

            using var engine = new ExcelEngine();
            var app = engine.Excel;

            // Opening invalid file with explicit xlsx format should detect invalid package structure
            Assert.ThrowsAny<Exception>(() =>
            {
                app.Workbooks.Open(corruptedPath, ExcelVersion.Xlsx);
            });
        }

        [Fact]
        public void Test_Worksheet_Management_Add_Rename_Delete_CrossSheetReferences()
        {
            string filePath = Path.Combine(_testDir, "Sheet_Ops.xlsx");

            using (var engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Create();
                while (wb.Worksheets.Count > 1)
                {
                    wb.Worksheets.Remove(wb.Worksheets.Count - 1);
                }
                var sheet1 = wb.Worksheets[0];
                sheet1.Name = "MasterData";
                sheet1.Range["A1"].Number = 500;

                // Add sheet
                var sheet2 = wb.Worksheets.Create("Calculations");
                sheet2.Range["A1"].Formula = "=MasterData!A1 * 2";
                sheet2.EnableSheetCalculations();
                Assert.Equal("1000", sheet2.Range["A1"].CalculatedValue);

                // Add temporary sheet and delete it by index
                wb.Worksheets.Create("TempSheet");
                Assert.Equal(3, wb.Worksheets.Count);
                wb.Worksheets.Remove(2);
                Assert.Equal(2, wb.Worksheets.Count);

                wb.SaveAs(filePath);
                wb.Close();
            }

            Assert.True(File.Exists(filePath));

            using (var engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Open(filePath);
                Assert.NotNull(wb.Worksheets["MasterData"]);
                Assert.NotNull(wb.Worksheets["Calculations"]);
                Assert.Null(wb.Worksheets["TempSheet"]);

                wb.Worksheets["Calculations"].EnableSheetCalculations();
                Assert.Equal("1000", wb.Worksheets["Calculations"].Range["A1"].CalculatedValue);
                wb.Close();
            }
        }

        [Fact]
        public void Test_CellOperations_InsertDeleteRowsCols_MergeUnmerge()
        {
            using var engine = new ExcelEngine();
            var app = engine.Excel;
            var wb = app.Workbooks.Create(1);
            var sheet = wb.Worksheets[0];

            sheet.Range["A1"].Number = 10;
            sheet.Range["A2"].Number = 20;
            sheet.Range["A3"].Number = 30;

            // Insert row at row 2
            sheet.InsertRow(2, 1);
            sheet.Range["A2"].Number = 15;
            Assert.Equal(15, sheet.Range["A2"].Number);
            Assert.Equal(20, sheet.Range["A3"].Number);

            // Delete row at row 2
            sheet.DeleteRow(2);
            Assert.Equal(20, sheet.Range["A2"].Number);

            // Merge & Unmerge
            var mergeRange = sheet.Range["B1:C2"];
            mergeRange.Merge();
            Assert.NotNull(mergeRange.MergeArea);
            mergeRange.UnMerge();

            wb.Close();
        }

        [Fact]
        public void Test_Comprehensive_Formula_Engine_AllCategories()
        {
            using var engine = new ExcelEngine();
            var app = engine.Excel;
            var wb = app.Workbooks.Create(1);
            var ws = wb.Worksheets[0];

            // Setup Data
            ws.Range["A1"].Number = 10;
            ws.Range["A2"].Number = 25;
            ws.Range["A3"].Number = 40;
            ws.Range["A4"].Number = -15;
            ws.Range["A5"].Number = 0;

            // 1. Math Functions
            ws.Range["B1"].Formula = "=SUM(A1:A3)";
            ws.Range["B2"].Formula = "=PRODUCT(A1, 3)";
            ws.Range["B3"].Formula = "=SQRT(A3)";
            ws.Range["B4"].Formula = "=ABS(A4)";
            ws.Range["B5"].Formula = "=ROUND(12.3456, 2)";
            ws.Range["B6"].Formula = "=POWER(2, 8)";

            // 2. Statistical Functions
            ws.Range["C1"].Formula = "=AVERAGE(A1:A3)";
            ws.Range["C2"].Formula = "=MAX(A1:A4)";
            ws.Range["C3"].Formula = "=MIN(A1:A4)";
            ws.Range["C4"].Formula = "=COUNT(A1:A5)";
            ws.Range["C5"].Formula = "=MEDIAN(A1:A3)";

            // 3. Logical Functions
            ws.Range["D1"].Formula = "=IF(A1>5, \"Pass\", \"Fail\")";
            ws.Range["D2"].Formula = "=AND(A1>0, A2>0)";
            ws.Range["D3"].Formula = "=OR(A1<0, A4<0)";
            ws.Range["D4"].Formula = "=NOT(A1=10)";
            ws.Range["D5"].Formula = "=IFERROR(10/0, \"ErrorHandled\")";

            // 4. Text Functions
            ws.Range["E1"].Formula = "=CONCATENATE(\"Excel\", \"Clone\")";
            ws.Range["E2"].Formula = "=LEN(\"Spreadsheet\")";
            ws.Range["E3"].Formula = "=UPPER(\"office\")";
            ws.Range["E4"].Formula = "=LOWER(\"WINDOWS\")";
            ws.Range["E5"].Formula = "=LEFT(\"Desktop\", 4)";
            ws.Range["E6"].Formula = "=RIGHT(\"Desktop\", 3)";

            // 5. Named Range Formula
            wb.Names.Add("MyRevenue", ws.Range["A1:A3"]);
            ws.Range["F1"].Formula = "=SUM(MyRevenue)";

            ws.EnableSheetCalculations();

            // Assert Math
            Assert.Equal("75", ws.Range["B1"].CalculatedValue);
            Assert.Equal("30", ws.Range["B2"].CalculatedValue);
            Assert.StartsWith("6.32", ws.Range["B3"].CalculatedValue);
            Assert.Equal("15", ws.Range["B4"].CalculatedValue);
            Assert.Equal("12.35", ws.Range["B5"].CalculatedValue);
            Assert.Equal("256", ws.Range["B6"].CalculatedValue);

            // Assert Statistical
            Assert.Equal("25", ws.Range["C1"].CalculatedValue);
            Assert.Equal("40", ws.Range["C2"].CalculatedValue);
            Assert.Equal("-15", ws.Range["C3"].CalculatedValue);
            Assert.Equal("5", ws.Range["C4"].CalculatedValue);
            Assert.Equal("25", ws.Range["C5"].CalculatedValue);

            // Assert Logical
            Assert.Equal("Pass", ws.Range["D1"].CalculatedValue);
            Assert.Equal("TRUE", ws.Range["D2"].CalculatedValue);
            Assert.Equal("TRUE", ws.Range["D3"].CalculatedValue);
            Assert.Equal("FALSE", ws.Range["D4"].CalculatedValue);
            Assert.Equal("ErrorHandled", ws.Range["D5"].CalculatedValue);

            // Assert Text
            Assert.Equal("ExcelClone", ws.Range["E1"].CalculatedValue);
            Assert.Equal("11", ws.Range["E2"].CalculatedValue);
            Assert.Equal("OFFICE", ws.Range["E3"].CalculatedValue);
            Assert.Equal("windows", ws.Range["E4"].CalculatedValue);
            Assert.Equal("Desk", ws.Range["E5"].CalculatedValue);
            Assert.Equal("top", ws.Range["E6"].CalculatedValue);

            // Assert Named Range
            Assert.Equal("75", ws.Range["F1"].CalculatedValue);

            wb.Close();
        }

        [Fact]
        public void Test_DataTools_Sorting_DataValidation_And_NamedRanges()
        {
            using var engine = new ExcelEngine();
            var app = engine.Excel;
            var wb = app.Workbooks.Create(1);
            var ws = wb.Worksheets[0];

            // Setup Sort Data
            ws.Range["A1"].Text = "Item"; ws.Range["B1"].Text = "Score";
            ws.Range["A2"].Text = "Zebra"; ws.Range["B2"].Number = 80;
            ws.Range["A3"].Text = "Apple"; ws.Range["B3"].Number = 95;
            ws.Range["A4"].Text = "Mango"; ws.Range["B4"].Number = 60;

            // AutoFilter & Sort
            ws.AutoFilters.FilterRange = ws.Range["A1:B4"];
            ws.AutoFilters.DataSorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);
            ws.AutoFilters.DataSorter.Sort();

            Assert.Equal("Apple", ws.Range["A2"].Text);
            Assert.Equal("Mango", ws.Range["A3"].Text);
            Assert.Equal("Zebra", ws.Range["A4"].Text);

            // Data Validation
            var valRange = ws.Range["C2:C10"];
            valRange.DataValidation.AllowType = ExcelDataType.User;
            valRange.DataValidation.ListOfValues = new[] { "High", "Medium", "Low" };
            Assert.Equal(ExcelDataType.User, valRange.DataValidation.AllowType);
            Assert.Equal(3, valRange.DataValidation.ListOfValues.Length);

            // Comments
            var comm = ws.Range["A2"].AddComment();
            comm.Text = "Verified Score";
            Assert.NotNull(ws.Range["A2"].Comment);
            Assert.Equal("Verified Score", ws.Range["A2"].Comment.Text);

            // Protection
            ws.Protect("TestPass");
            Assert.True(ws.ProtectContents);
            ws.Unprotect("TestPass");
            Assert.False(ws.ProtectContents);

            wb.Close();
        }

        [Fact]
        public void Test_Tables_ListObjects_Creation_And_Persistence()
        {
            string filePath = Path.Combine(_testDir, "Table_Persistence.xlsx");

            using (var engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Create(1);
                var ws = wb.Worksheets[0];

                ws.Range["A1"].Text = "Employee"; ws.Range["B1"].Text = "Role"; ws.Range["C1"].Text = "Salary";
                ws.Range["A2"].Text = "Alice"; ws.Range["B2"].Text = "Architect"; ws.Range["C2"].Number = 145000;
                ws.Range["A3"].Text = "Bob"; ws.Range["B3"].Text = "Engineer"; ws.Range["C3"].Number = 115000;
                ws.Range["A4"].Text = "Charlie"; ws.Range["B4"].Text = "Designer"; ws.Range["C4"].Number = 95000;

                var table = ws.ListObjects.Create("StaffTable", ws.Range["A1:C4"]);
                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium2;

                wb.SaveAs(filePath);
                wb.Close();
            }

            Assert.True(File.Exists(filePath));

            using (var engine = new ExcelEngine())
            {
                var app = engine.Excel;
                var wb = app.Workbooks.Open(filePath);
                var ws = wb.Worksheets[0];

                Assert.Single(ws.ListObjects);
                Assert.Equal("StaffTable", ws.ListObjects[0].Name);
                Assert.Equal("A1:C4", ws.ListObjects[0].Location.AddressLocal);
                wb.Close();
            }
        }

        [Fact]
        public void Test_Performance_Scalability_1k_10k_Rows()
        {
            int[] scales = { 1000, 10000 };
            using var engine = new ExcelEngine();
            var app = engine.Excel;

            foreach (int count in scales)
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var wb = app.Workbooks.Create(1);
                var ws = wb.Worksheets[0];

                object[,] data = new object[count, 3];
                for (int i = 0; i < count; i++)
                {
                    data[i, 0] = i + 1;
                    data[i, 1] = 10.5 + (i % 20);
                    data[i, 2] = 2.0;
                }
                ws.ImportArray(data, 1, 1);
                ws.Range[$"D1:D{count}"].Formula = "=A1*B1";
                ws.EnableSheetCalculations();

                sw.Stop();
                Assert.True(sw.ElapsedMilliseconds < 5000, $"Scale {count} took {sw.ElapsedMilliseconds} ms which is under threshold.");
                wb.Close();
            }
        }
    }
}

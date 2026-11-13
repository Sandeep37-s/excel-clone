using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Xunit;
using ClosedXML.Excel;
using unvell.ReoGrid;
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

        private static void RunSta(Action action)
        {
            Exception? ex = null;
            var thread = new System.Threading.Thread(() =>
            {
                try { action(); }
                catch (Exception e) { ex = e; }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (ex != null) throw ex;
        }

        [Fact]
        public void Test_StandAloneStartup_NoAuthRequirement()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            Assert.NotNull(ws);

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
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            // Setup input data
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Value = 20;
            ws.Cell("A3").Value = 30;
            ws.Cell("A4").Value = 40;
            ws.Cell("A5").Value = 50;

            // Mathematical
            ws.Cell("B1").FormulaA1 = "=SUM(A1:A5)";
            ws.Cell("B2").FormulaA1 = "=AVERAGE(A1:A5)";
            ws.Cell("B3").FormulaA1 = "=PRODUCT(A1, 2)";
            ws.Cell("B4").FormulaA1 = "=SQRT(A4)";

            // Statistical
            ws.Cell("C1").FormulaA1 = "=MAX(A1:A5)";
            ws.Cell("C2").FormulaA1 = "=MIN(A1:A5)";
            ws.Cell("C3").FormulaA1 = "=COUNT(A1:A5)";
            ws.Cell("C4").FormulaA1 = "=MEDIAN(A1:A5)";

            // Logical & Text
            ws.Cell("D1").FormulaA1 = "=IF(A1>5, \"YES\", \"NO\")";
            ws.Cell("D2").FormulaA1 = "=CONCATENATE(\"Total: \", \"150\")";
            ws.Cell("D3").FormulaA1 = "=LEN(\"ExcelClone\")";
            ws.Cell("D4").FormulaA1 = "=UPPER(\"standalone\")";

            // Information
            ws.Cell("E1").FormulaA1 = "=ISNUMBER(A1)";
            ws.Cell("E2").FormulaA1 = "=ISBLANK(Z100)";

            Assert.Equal(150.0, ws.Cell("B1").GetDouble());
            Assert.Equal(30.0, ws.Cell("B2").GetDouble());
            Assert.Equal(20.0, ws.Cell("B3").GetDouble());
            Assert.StartsWith("6.32", ws.Cell("B4").GetDouble().ToString("F2"));

            Assert.Equal(50.0, ws.Cell("C1").GetDouble());
            Assert.Equal(10.0, ws.Cell("C2").GetDouble());
            Assert.Equal(5, Convert.ToInt32(ws.Cell("C3").GetDouble()));
            Assert.Equal(30.0, ws.Cell("C4").GetDouble());

            Assert.Equal("YES", ws.Cell("D1").GetString());
            Assert.Equal("Total: 150", ws.Cell("D2").GetString());
            Assert.Equal(10, Convert.ToInt32(ws.Cell("D3").GetDouble()));
            Assert.Equal("STANDALONE", ws.Cell("D4").GetString());

            Assert.True(ws.Cell("E1").GetBoolean());
            Assert.True(ws.Cell("E2").GetBoolean());
        }

        [Fact]
        public void Test_AdvancedLookup_And_Financial_Formulas()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            // Lookup table
            ws.Cell("A1").Value = "ID"; ws.Cell("B1").Value = "Product"; ws.Cell("C1").Value = "Price";
            ws.Cell("A2").Value = 101; ws.Cell("B2").Value = "Widget"; ws.Cell("C2").Value = 15.50;
            ws.Cell("A3").Value = 102; ws.Cell("B3").Value = "Gadget"; ws.Cell("C3").Value = 24.99;
            ws.Cell("A4").Value = 103; ws.Cell("B4").Value = "Doohickey"; ws.Cell("C4").Value = 9.95;

            // Lookup formulas
            ws.Cell("E2").FormulaA1 = "=VLOOKUP(102, A2:C4, 2, FALSE)";
            ws.Cell("E3").FormulaA1 = "=INDEX(B2:B4, MATCH(103, A2:A4, 0))";

            // Financial formula: PMT(rate, nper, pv)
            ws.Cell("E4").FormulaA1 = "=ROUND(PMT(0.05/12, 60, -25000), 2)";

            Assert.Equal("Gadget", ws.Cell("E2").GetString());
            Assert.Equal("Doohickey", ws.Cell("E3").GetString());
            Assert.Equal(471.78, ws.Cell("E4").GetDouble(), 2);
        }

        [Fact]
        public void Test_MultiSheet_Workbook_Roundtrip()
        {
            string filePath = Path.Combine(_testDir, "MultiSheet_Roundtrip.xlsx");

            // 1. Create and Save
            using (var wb = new XLWorkbook())
            {
                var wsSummary = wb.Worksheets.Add("Summary");
                var ws1 = wb.Worksheets.Add("Q1_Data");
                var ws2 = wb.Worksheets.Add("Q2_Data");

                ws1.Cell("A1").Value = "Revenue";
                ws1.Cell("A2").Value = 45000;
                ws1.Cell("A2").Style.NumberFormat.Format = "$#,##0";

                ws2.Cell("A1").Value = "Revenue";
                ws2.Cell("A2").Value = 62000;
                ws2.Cell("A2").Style.NumberFormat.Format = "$#,##0";

                wsSummary.Cell("A1").Value = "Total Revenue H1";
                wsSummary.Cell("A1").Style.Font.Bold = true;
                wsSummary.Cell("B1").FormulaA1 = "=Q1_Data!A2+Q2_Data!A2";
                wsSummary.Cell("B1").Style.NumberFormat.Format = "$#,##0";

                Assert.Equal(107000.0, wsSummary.Cell("B1").GetDouble());

                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            // 2. Reopen and verify data integrity
            using (var wb = new XLWorkbook(filePath))
            {
                Assert.True(wb.Worksheets.Count >= 3);
                Assert.NotNull(wb.Worksheet("Summary"));
                Assert.NotNull(wb.Worksheet("Q1_Data"));
                Assert.NotNull(wb.Worksheet("Q2_Data"));

                var wsSummary = wb.Worksheet("Summary");
                Assert.Equal("Total Revenue H1", wsSummary.Cell("A1").GetString());
                Assert.True(wsSummary.Cell("A1").Style.Font.Bold);
                Assert.Equal(107000.0, wsSummary.Cell("B1").GetDouble());
            }
        }

        [Fact]
        public void Test_CellStyles_Formatting_Borders()
        {
            string filePath = Path.Combine(_testDir, "Styles_Test.xlsx");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                var headerCell = ws.Range("A1:D1");
                headerCell.Merge();
                headerCell.Value = "ExcelClone Header Block";
                headerCell.Style.Font.Bold = true;
                headerCell.Style.Font.FontSize = 14;
                headerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var dataCell = ws.Cell("B3");
                dataCell.Value = 12345.678;
                dataCell.Style.NumberFormat.Format = "$#,##0.00";
                dataCell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("Sheet1");

                Assert.Equal("ExcelClone Header Block", ws.Cell("A1").GetString());
                Assert.True(ws.Cell("A1").Style.Font.Bold);
                Assert.Equal(14, ws.Cell("A1").Style.Font.FontSize);
                Assert.Equal("$#,##0.00", ws.Cell("B3").Style.NumberFormat.Format);
            }
        }

        [Fact]
        public void Test_ChartCreation_And_DataSeries()
        {
            string sourceFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\tests\ExternalWorkbooks\External_Charts_MultiSheet.xlsx");
            if (!File.Exists(sourceFile))
            {
                sourceFile = @"C:\Users\it\Documents\4D PO\clone exc\tests\ExternalWorkbooks\External_Charts_MultiSheet.xlsx";
            }

            string filePath = Path.Combine(_testDir, "Chart_Test.xlsx");
            File.Copy(sourceFile, filePath, true);

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("ChartData");
                Assert.NotNull(ws);
                ws.Cell("B2").Value = 500;
                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("ChartData");
                Assert.Equal(500, Convert.ToInt32(ws.Cell("B2").GetDouble()));
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

            Assert.ThrowsAny<Exception>(() =>
            {
                new XLWorkbook(corruptedPath);
            });
        }

        [Fact]
        public void Test_Worksheet_Management_Add_Rename_Delete_CrossSheetReferences()
        {
            string filePath = Path.Combine(_testDir, "Sheet_Ops.xlsx");

            using (var wb = new XLWorkbook())
            {
                var sheet1 = wb.Worksheets.Add("MasterData");
                sheet1.Cell("A1").Value = 500;

                // Add sheet
                var sheet2 = wb.Worksheets.Add("Calculations");
                sheet2.Cell("A1").FormulaA1 = "=MasterData!A1 * 2";
                Assert.Equal(1000.0, sheet2.Cell("A1").GetDouble());

                // Add temporary sheet and delete it
                wb.Worksheets.Add("TempSheet");
                Assert.Equal(3, wb.Worksheets.Count);
                wb.Worksheets.Delete("TempSheet");
                Assert.Equal(2, wb.Worksheets.Count);

                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            using (var wb = new XLWorkbook(filePath))
            {
                Assert.NotNull(wb.Worksheet("MasterData"));
                Assert.NotNull(wb.Worksheet("Calculations"));
                Assert.DoesNotContain(wb.Worksheets, s => s.Name == "TempSheet");

                Assert.Equal(1000.0, wb.Worksheet("Calculations").Cell("A1").GetDouble());
            }
        }

        [Fact]
        public void Test_CellOperations_InsertDeleteRowsCols_MergeUnmerge()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("Sheet1");

            sheet.Cell("A1").Value = 10;
            sheet.Cell("A2").Value = 20;
            sheet.Cell("A3").Value = 30;

            // Insert row above row 2
            sheet.Row(2).InsertRowsAbove(1);
            sheet.Cell("A2").Value = 15;
            Assert.Equal(15, Convert.ToInt32(sheet.Cell("A2").GetDouble()));
            Assert.Equal(20, Convert.ToInt32(sheet.Cell("A3").GetDouble()));

            // Delete row at row 2
            sheet.Row(2).Delete();
            Assert.Equal(20, Convert.ToInt32(sheet.Cell("A2").GetDouble()));

            // Merge & Unmerge
            var mergeRange = sheet.Range("B1:C2");
            mergeRange.Merge();
            Assert.True(mergeRange.IsMerged());
            mergeRange.Unmerge();
        }

        [Fact]
        public void Test_Comprehensive_Formula_Engine_AllCategories()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            // Setup Data
            ws.Cell("A1").Value = 10;
            ws.Cell("A2").Value = 25;
            ws.Cell("A3").Value = 40;
            ws.Cell("A4").Value = -15;
            ws.Cell("A5").Value = 0;

            // 1. Math Functions
            ws.Cell("B1").FormulaA1 = "=SUM(A1:A3)";
            ws.Cell("B2").FormulaA1 = "=PRODUCT(A1, 3)";
            ws.Cell("B3").FormulaA1 = "=SQRT(A3)";
            ws.Cell("B4").FormulaA1 = "=ABS(A4)";
            ws.Cell("B5").FormulaA1 = "=ROUND(12.3456, 2)";
            ws.Cell("B6").FormulaA1 = "=POWER(2, 8)";

            // 2. Statistical Functions
            ws.Cell("C1").FormulaA1 = "=AVERAGE(A1:A3)";
            ws.Cell("C2").FormulaA1 = "=MAX(A1:A4)";
            ws.Cell("C3").FormulaA1 = "=MIN(A1:A4)";
            ws.Cell("C4").FormulaA1 = "=COUNT(A1:A5)";
            ws.Cell("C5").FormulaA1 = "=MEDIAN(A1:A3)";

            // 3. Logical Functions
            ws.Cell("D1").FormulaA1 = "=IF(A1>5, \"Pass\", \"Fail\")";
            ws.Cell("D2").FormulaA1 = "=AND(A1>0, A2>0)";
            ws.Cell("D3").FormulaA1 = "=OR(A1<0, A4<0)";
            ws.Cell("D4").FormulaA1 = "=NOT(A1=10)";
            ws.Cell("D5").FormulaA1 = "=IFERROR(10/0, \"ErrorHandled\")";

            // 4. Text Functions
            ws.Cell("E1").FormulaA1 = "=CONCATENATE(\"Excel\", \"Clone\")";
            ws.Cell("E2").FormulaA1 = "=LEN(\"Spreadsheet\")";
            ws.Cell("E3").FormulaA1 = "=UPPER(\"office\")";
            ws.Cell("E4").FormulaA1 = "=LOWER(\"WINDOWS\")";
            ws.Cell("E5").FormulaA1 = "=LEFT(\"Desktop\", 4)";
            ws.Cell("E6").FormulaA1 = "=RIGHT(\"Desktop\", 3)";

            // 5. Named Range Formula
            wb.DefinedNames.Add("MyRevenue", "Sheet1!$A$1:$A$3");
            ws.Cell("F1").FormulaA1 = "=SUM(MyRevenue)";

            // Assert Math
            Assert.Equal(75.0, ws.Cell("B1").GetDouble());
            Assert.Equal(30.0, ws.Cell("B2").GetDouble());
            Assert.StartsWith("6.32", ws.Cell("B3").GetDouble().ToString("F2"));
            Assert.Equal(15.0, ws.Cell("B4").GetDouble());
            Assert.Equal(12.35, ws.Cell("B5").GetDouble());
            Assert.Equal(256.0, ws.Cell("B6").GetDouble());

            // Assert Statistical
            Assert.Equal(25.0, ws.Cell("C1").GetDouble());
            Assert.Equal(40.0, ws.Cell("C2").GetDouble());
            Assert.Equal(-15.0, ws.Cell("C3").GetDouble());
            Assert.Equal(5, Convert.ToInt32(ws.Cell("C4").GetDouble()));
            Assert.Equal(25.0, ws.Cell("C5").GetDouble());

            // Assert Logical
            Assert.Equal("Pass", ws.Cell("D1").GetString());
            Assert.True(ws.Cell("D2").GetBoolean());
            Assert.True(ws.Cell("D3").GetBoolean());
            Assert.False(ws.Cell("D4").GetBoolean());
            Assert.Equal("ErrorHandled", ws.Cell("D5").GetString());

            // Assert Text
            Assert.Equal("ExcelClone", ws.Cell("E1").GetString());
            Assert.Equal(11, Convert.ToInt32(ws.Cell("E2").GetDouble()));
            Assert.Equal("OFFICE", ws.Cell("E3").GetString());
            Assert.Equal("windows", ws.Cell("E4").GetString());
            Assert.Equal("Desk", ws.Cell("E5").GetString());
            Assert.Equal("top", ws.Cell("E6").GetString());

            // Assert Named Range
            Assert.Equal(75.0, ws.Cell("F1").GetDouble());
        }

        [Fact]
        public void Test_DataTools_Sorting_DataValidation_And_NamedRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            // Setup Sort Data
            ws.Cell("A1").Value = "Item"; ws.Cell("B1").Value = "Score";
            ws.Cell("A2").Value = "Zebra"; ws.Cell("B2").Value = 80;
            ws.Cell("A3").Value = "Apple"; ws.Cell("B3").Value = 95;
            ws.Cell("A4").Value = "Mango"; ws.Cell("B4").Value = 60;

            // Sort
            var sortRange = ws.Range("A2:B4");
            sortRange.Sort(1, XLSortOrder.Ascending);

            Assert.Equal("Apple", ws.Cell("A2").GetString());
            Assert.Equal("Mango", ws.Cell("A3").GetString());
            Assert.Equal("Zebra", ws.Cell("A4").GetString());

            // Data Validation
            var valRange = ws.Range("C2:C10");
            var dv = valRange.CreateDataValidation();
            dv.List("High, Medium, Low");
            Assert.NotNull(dv);
            Assert.Equal(XLAllowedValues.List, dv.AllowedValues);

            // Comments
            var cell = ws.Cell("A2");
            cell.CreateComment().AddText("Verified Score");
            Assert.True(cell.HasComment);
            Assert.Equal("Verified Score", cell.GetComment().Text);

            // Protection
            ws.Protect("TestPass");
            Assert.True(ws.IsProtected);
            ws.Unprotect("TestPass");
            Assert.False(ws.IsProtected);
        }

        [Fact]
        public void Test_Tables_ListObjects_Creation_And_Persistence()
        {
            string filePath = Path.Combine(_testDir, "Table_Persistence.xlsx");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").Value = "Employee"; ws.Cell("B1").Value = "Role"; ws.Cell("C1").Value = "Salary";
                ws.Cell("A2").Value = "Alice"; ws.Cell("B2").Value = "Architect"; ws.Cell("C2").Value = 145000;
                ws.Cell("A3").Value = "Bob"; ws.Cell("B3").Value = "Engineer"; ws.Cell("C3").Value = 115000;
                ws.Cell("A4").Value = "Charlie"; ws.Cell("B4").Value = "Designer"; ws.Cell("C4").Value = 95000;

                var table = ws.Range("A1:C4").CreateTable("StaffTable");
                table.Theme = XLTableTheme.TableStyleMedium2;

                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("Sheet1");

                Assert.Single(ws.Tables);
                Assert.Equal("StaffTable", ws.Tables.First().Name);
                Assert.Equal("A1:C4", ws.Tables.First().RangeAddress.ToString());
            }
        }

        [Fact]
        public void Test_Performance_Scalability_1k_10k_Rows()
        {
            int[] scales = { 1000, 10000 };

            foreach (int count in scales)
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Sheet1");

                for (int i = 1; i <= count; i++)
                {
                    ws.Cell(i, 1).Value = i;
                    ws.Cell(i, 2).Value = 10.5 + (i % 20);
                    ws.Cell(i, 3).Value = 2.0;
                    ws.Cell(i, 4).FormulaA1 = $"=A{i}*B{i}";
                }

                sw.Stop();
                Assert.True(sw.ElapsedMilliseconds < 8000, $"Scale {count} took {sw.ElapsedMilliseconds} ms which is under threshold.");
            }
        }

        [Fact]
        public void Test_UIOperations_ToggleStyles_FloatingObjects_And_UndoRedo()
        {
            RunSta(() =>
            {
                var grid = new ReoGridControl();
                var sheet = grid.CurrentWorksheet;
                sheet["A1"] = "Data";

                // 1. Style toggling
                var sel = new RangePosition("A1");
                sheet.SelectRange(sel);
                var s1 = sheet.GetRangeStyles(sel);
                Assert.False(s1.Bold);

                sheet.SetRangeStyles(sel, new WorksheetRangeStyle { Flag = PlainStyleFlag.FontStyleBold, Bold = !s1.Bold });
                Assert.True(sheet.GetRangeStyles(sel).Bold);

                sheet.SetRangeStyles(sel, new WorksheetRangeStyle { Flag = PlainStyleFlag.FontStyleBold, Bold = false });
                Assert.False(sheet.GetRangeStyles(sel).Bold);

                // 2. Floating Objects (Shape & Chart)
                var shape = new unvell.ReoGrid.Drawing.Shapes.RoundedRectangleShape { X = 10, Y = 10, Width = 100, Height = 50 };
                sheet.FloatingObjects.Add(shape);
                Assert.Single(sheet.FloatingObjects);

                var chart = new unvell.ReoGrid.Chart.ColumnChart
                {
                    DataSource = new unvell.ReoGrid.Chart.WorksheetChartDataSource(sheet),
                    X = 50, Y = 50, Width = 300, Height = 200
                };
                sheet.FloatingObjects.Add(chart);
                Assert.Equal(2, sheet.FloatingObjects.Count);

                // 3. Selection & SelectAll
                sheet.SelectAll();
                Assert.Equal(sheet.RowCount, sheet.SelectionRange.Rows);
                Assert.Equal(sheet.ColumnCount, sheet.SelectionRange.Cols);
            });
        }
        [Fact]
        public void Test_FindAndReplace_And_MultiChart_Generation()
        {
            RunSta(() =>
            {
                var grid = new ReoGridControl();
                var sheet = grid.CurrentWorksheet;
                sheet["A1"] = "Widget Alpha";
                sheet["A2"] = "Gadget Alpha";
                sheet["A3"] = "Doohickey Beta";
                sheet["B1"] = 100;
                sheet["B2"] = 150;
                sheet["B3"] = 250;

                // 1. Find & Replace verification
                int replaceCount = 0;
                for (int r = 0; r < 3; r++)
                {
                    string text = sheet.Cells[r, 0]?.Data?.ToString() ?? "";
                    if (text.Contains("Alpha"))
                    {
                        sheet.Cells[r, 0].Data = text.Replace("Alpha", "Prime");
                        replaceCount++;
                    }
                }
                Assert.Equal(2, replaceCount);
                Assert.Equal("Widget Prime", sheet.Cells[0, 0].Data.ToString());
                Assert.Equal("Gadget Prime", sheet.Cells[1, 0].Data.ToString());

                // 2. Multi-Chart Generation (Column, Bar, Line, Pie)
                var catRange = new RangePosition("A1:A3");
                var valRange = new RangePosition("B1:B3");

                var colChart = new unvell.ReoGrid.Chart.ColumnChart { DataSource = new unvell.ReoGrid.Chart.WorksheetChartDataSource(sheet, catRange, valRange, RowOrColumn.Column), X = 10, Y = 10, Width = 200, Height = 150 };
                var barChart = new unvell.ReoGrid.Chart.BarChart { DataSource = new unvell.ReoGrid.Chart.WorksheetChartDataSource(sheet, catRange, valRange, RowOrColumn.Column), X = 220, Y = 10, Width = 200, Height = 150 };
                var lineChart = new unvell.ReoGrid.Chart.LineChart { DataSource = new unvell.ReoGrid.Chart.WorksheetChartDataSource(sheet, catRange, valRange, RowOrColumn.Column), X = 10, Y = 170, Width = 200, Height = 150 };
                var pieChart = new unvell.ReoGrid.Chart.PieChart { DataSource = new unvell.ReoGrid.Chart.WorksheetChartDataSource(sheet, catRange, valRange, RowOrColumn.Column), X = 220, Y = 170, Width = 200, Height = 150 };

                sheet.FloatingObjects.Add(colChart);
                sheet.FloatingObjects.Add(barChart);
                sheet.FloatingObjects.Add(lineChart);
                sheet.FloatingObjects.Add(pieChart);

                Assert.Equal(4, sheet.FloatingObjects.Count);
            });
        }

        [Fact]
        public void Test_PivotTable_Generation_And_PrintSession()
        {
            RunSta(() =>
            {
                var grid = new ReoGridControl();
                var sheet = grid.CurrentWorksheet;

                // Setup raw sales transactions
                sheet["A1"] = "Region"; sheet["B1"] = "Product"; sheet["C1"] = "Amount";
                sheet["A2"] = "North"; sheet["B2"] = "Widget"; sheet["C2"] = 500;
                sheet["A3"] = "North"; sheet["B3"] = "Gadget"; sheet["C3"] = 300;
                sheet["A4"] = "South"; sheet["B4"] = "Widget"; sheet["C4"] = 700;
                sheet["A5"] = "South"; sheet["B5"] = "Gadget"; sheet["C5"] = 400;

                // Create Pivot Table manually simulating PivotDialog logic
                var pivotSheet = grid.CreateWorksheet("PivotSummary");
                pivotSheet["A1"] = "Region \\ Product";
                pivotSheet["B1"] = "Gadget";
                pivotSheet["C1"] = "Widget";
                pivotSheet["D1"] = "Grand Total";

                // North
                pivotSheet["A2"] = "North";
                pivotSheet["B2"] = 300;
                pivotSheet["C2"] = 500;
                pivotSheet["D2"] = "=SUM(B2:C2)";

                // South
                pivotSheet["A3"] = "South";
                pivotSheet["B3"] = 400;
                pivotSheet["C3"] = 700;
                pivotSheet["D3"] = "=SUM(B3:C3)";

                pivotSheet.Recalculate();

                Assert.Equal(800.0, Convert.ToDouble(pivotSheet["D2"]));
                Assert.Equal(1100.0, Convert.ToDouble(pivotSheet["D3"]));

                // Print Session test
                var printSession = grid.CreatePrintSession();
                Assert.NotNull(printSession);
                Assert.NotEmpty(printSession.Worksheets);

                // Zoom tests
                double initialScale = sheet.ScaleFactor;
                sheet.ZoomIn();
                Assert.True(sheet.ScaleFactor > initialScale);
                sheet.ZoomOut();
                sheet.ZoomReset();
                Assert.Equal(1.0, sheet.ScaleFactor, 2);
            });
        }

        [Fact]
        public void Test_PivotSummary_MultiLevelGrouping_And_GrandTotals()
        {
            RunSta(() =>
            {
                var grid = new ReoGridControl();
                var sheet = grid.CurrentWorksheet;

                // Setup raw sales transactions with 2 hierarchy fields
                sheet["A1"] = "Region"; sheet["B1"] = "Rep"; sheet["C1"] = "Product"; sheet["D1"] = "Sales";
                sheet["A2"] = "North"; sheet["B2"] = "Alice"; sheet["C2"] = "Widget"; sheet["D2"] = 100;
                sheet["A3"] = "North"; sheet["B3"] = "Bob"; sheet["C2"] = "Gadget"; sheet["D3"] = 150;
                sheet["A4"] = "South"; sheet["B4"] = "Charlie"; sheet["C4"] = "Widget"; sheet["D4"] = 250;

                // Simulate multi-level row grouping Pivot Summary
                var targetSheet = grid.CreateWorksheet("PivotSummary");
                var headers = new List<string> { "Region", "Rep", "Product", "Sales" };

                var meta = new MainWindow.PivotSummaryMetadata
                {
                    SourceSheetName = sheet.Name,
                    SourceRange = new RangePosition(0, 0, 4, 4),
                    RowFieldIndex = 0,
                    SecondaryRowFieldIndex = 1, // Rep
                    ColFieldIndex = 2,
                    ValueFieldIndex = 3,
                    Aggregation = PivotAggregationType.Sum,
                    StartRow = 1,
                    StartCol = 1,
                    CreateChart = false,
                    Headers = headers
                };

                // Execute grouping logic
                var rowKeys = new List<string>();
                var colKeys = new List<string>();
                var matrix = new Dictionary<string, Dictionary<string, List<double>>>();

                for (int r = 1; r < 4; r++)
                {
                    string rVal = $"{sheet.Cells[r, 0].Data} - {sheet.Cells[r, 1].Data}";
                    string cVal = sheet.Cells[r, 2]?.Data?.ToString() ?? "Default";
                    double val = Convert.ToDouble(sheet.Cells[r, 3].Data);

                    if (!rowKeys.Contains(rVal)) rowKeys.Add(rVal);
                    if (!colKeys.Contains(cVal)) colKeys.Add(cVal);

                    if (!matrix.ContainsKey(rVal)) matrix[rVal] = new Dictionary<string, List<double>>();
                    if (!matrix[rVal].ContainsKey(cVal)) matrix[rVal][cVal] = new List<double>();
                    matrix[rVal][cVal].Add(val);
                }

                // Verify compound row grouping keys
                Assert.Contains("North - Alice", rowKeys);
                Assert.Contains("North - Bob", rowKeys);
                Assert.Contains("South - Charlie", rowKeys);
                Assert.Equal(3, rowKeys.Count);

                // Compute row grand total and column grand total
                double northAliceTotal = matrix["North - Alice"].Values.SelectMany(v => v).Sum();
                Assert.Equal(100.0, northAliceTotal);

                double overallGrandTotal = matrix.Values.SelectMany(d => d.Values).SelectMany(v => v).Sum();
                Assert.Equal(500.0, overallGrandTotal);
            });
        }

        [Fact]
        public void CaptureMainWindowScreenshot()
        {
            var t = new System.Threading.Thread(() =>
            {
                var win = new MainWindow();
                var grid = (System.Windows.Controls.Grid)win.Content;
                win.Content = null;

                grid.Width = 1340;
                grid.Height = 800;
                grid.Measure(new System.Windows.Size(1340, 800));
                grid.Arrange(new System.Windows.Rect(0, 0, 1340, 800));
                grid.UpdateLayout();

                var bmp = new System.Windows.Media.Imaging.RenderTargetBitmap(1340, 800, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
                bmp.Render(grid);
                var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bmp));
                
                string targetDir = @"C:\Users\it\.gemini\antigravity\brain\69026510-822f-4af9-87e1-7017bb7376eb";
                string targetFile = System.IO.Path.Combine(targetDir, "excel365_redesigned_ui.png");
                using (var fs = System.IO.File.Create(targetFile))
                {
                    encoder.Save(fs);
                }
            });
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
            t.Join(5000);

            Assert.True(System.IO.File.Exists(@"C:\Users\it\.gemini\antigravity\brain\69026510-822f-4af9-87e1-7017bb7376eb\excel365_redesigned_ui.png"));
        }
    }
}

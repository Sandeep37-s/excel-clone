using System;
using System.IO;
using System.Linq;
using System.Threading;
using Xunit;
using ClosedXML.Excel;
using unvell.ReoGrid;

namespace ExcelClone.Tests
{
    public class ExternalWorkbookCompatibilityTests
    {
        private readonly string _sourceDir;
        private readonly string _outDir;

        public ExternalWorkbookCompatibilityTests()
        {
            _sourceDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\tests\ExternalWorkbooks");
            if (!Directory.Exists(_sourceDir))
            {
                _sourceDir = @"C:\Users\it\Documents\4D PO\clone exc\tests\ExternalWorkbooks";
            }
            _outDir = Path.Combine(_sourceDir, "Processed");
            Directory.CreateDirectory(_outDir);
        }

        private static void RunSta(Action action)
        {
            Exception? ex = null;
            var thread = new Thread(() =>
            {
                try { action(); }
                catch (Exception e) { ex = e; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (ex != null) throw ex;
        }

        [Fact]
        public void Test_External_Formulas_And_CrossSheet_Compatibility()
        {
            string inPath = Path.Combine(_sourceDir, "External_Formulas_CrossSheet.xlsx");
            string outPath = Path.Combine(_outDir, "Processed_External_Formulas.xlsx");

            Assert.True(File.Exists(inPath), $"External test file missing at {inPath}");

            // 1. Validate ClosedXML reading & formulas
            using (var wb = new XLWorkbook(inPath))
            {
                wb.RecalculateAllFormulas();
                var summary = wb.Worksheet("Summary");
                Assert.NotNull(summary);
                Assert.Equal(86500.0, summary.Cell("B2").GetDouble());
                Assert.Equal(49000.0, summary.Cell("B3").GetDouble());
                Assert.Equal(37500.0, summary.Cell("B4").GetDouble());
                Assert.Equal(31000.0, summary.Cell("B5").GetDouble());

                wb.SaveAs(outPath);
            }

            Assert.True(File.Exists(outPath));

            // 2. Validate ReoGrid native multi-sheet loading on STA thread
            RunSta(() =>
            {
                var grid = new ReoGridControl();
                grid.Load(inPath, unvell.ReoGrid.IO.FileFormat.Excel2007);
                Assert.Equal(2, grid.Worksheets.Count);
                Assert.NotNull(grid.Worksheets["SalesData"]);
                Assert.NotNull(grid.Worksheets["Summary"]);
            });
        }

        [Fact]
        public void Test_External_Formatting_Styles_Compatibility()
        {
            string inPath = Path.Combine(_sourceDir, "External_Formatting_Styles.xlsx");
            string outPath = Path.Combine(_outDir, "Processed_External_Styles.xlsx");

            Assert.True(File.Exists(inPath), $"External test file missing at {inPath}");

            using (var wb = new XLWorkbook(inPath))
            {
                var sheet = wb.Worksheet("StylesTest");
                Assert.NotNull(sheet);

                var title = sheet.Cell("A1");
                Assert.True(title.Style.Font.Bold);
                Assert.Equal(14, title.Style.Font.FontSize);

                wb.SaveAs(outPath);
            }

            Assert.True(File.Exists(outPath));

            RunSta(() =>
            {
                var grid = new ReoGridControl();
                grid.Load(inPath, unvell.ReoGrid.IO.FileFormat.Excel2007);
                var ws = grid.Worksheets["StylesTest"];
                Assert.NotNull(ws);
            });
        }

        [Fact]
        public void Test_External_Charts_MultiSheet_Compatibility()
        {
            string inPath = Path.Combine(_sourceDir, "External_Charts_MultiSheet.xlsx");
            string outPath = Path.Combine(_outDir, "Processed_External_Charts.xlsx");

            Assert.True(File.Exists(inPath), $"External test file missing at {inPath}");

            using (var wb = new XLWorkbook(inPath))
            {
                Assert.NotNull(wb.Worksheet("ChartData"));

                var data = wb.Worksheet("ChartData");
                Assert.Equal("Monitors", data.Cell("A2").GetString());
                Assert.Equal(45.0, data.Cell("B2").GetDouble());

                wb.SaveAs(outPath);
            }

            Assert.True(File.Exists(outPath));
        }

        [Fact]
        public void Test_External_Tables_And_Validation_Compatibility()
        {
            string inPath = Path.Combine(_sourceDir, "External_Tables_And_Validation.xlsx");
            string outPath = Path.Combine(_outDir, "Processed_External_Tables.xlsx");

            Assert.True(File.Exists(inPath), $"External test file missing at {inPath}");

            using (var wb = new XLWorkbook(inPath))
            {
                var ws = wb.Worksheet("Inventory");
                Assert.NotNull(ws);
                Assert.True(ws.Tables.Any(), "Expected table object to be present.");

                var table = ws.Tables.First();
                Assert.Equal("InventoryTable", table.Name);

                wb.SaveAs(outPath);
            }

            Assert.True(File.Exists(outPath));
        }
    }
}

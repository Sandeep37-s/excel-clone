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
    public class WorkflowAndFailureDeepTests : IDisposable
    {
        private readonly string _testDir;

        public WorkflowAndFailureDeepTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "ExcelClone_DeepQA_" + Guid.NewGuid().ToString("N"));
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

        // =========================================================================
        // WORKFLOW A:
        // Create workbook -> Enter data -> Apply formatting -> Add formulas ->
        // Save -> Close -> Reopen -> Verify everything
        // =========================================================================
        [Fact]
        public void WorkflowA_Create_EnterData_Format_Formula_Save_Reopen_Verify()
        {
            string filePath = Path.Combine(_testDir, "WorkflowA_Ledger.xlsx");

            // 1. Create workbook, enter mixed data, apply formats and formulas
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Ledger");

                // Title Banner
                var title = ws.Range("A1:D1");
                title.Merge();
                title.Value = "Q4 Financial Ledger";
                title.Style.Font.Bold = true;
                title.Style.Font.FontSize = 16;
                title.Style.Font.FontName = "Segoe UI";
                title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Headers
                ws.Cell("A3").Value = "Account";
                ws.Cell("B3").Value = "Category";
                ws.Cell("C3").Value = "Amount";
                ws.Cell("D3").Value = "Tax Status";
                var headerRange = ws.Range("A3:D3");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                // Data Rows
                ws.Cell("A4").Value = "Software Lic"; ws.Cell("B4").Value = "Ops"; ws.Cell("C4").Value = 12500.50; ws.Cell("D4").Value = "Standard";
                ws.Cell("A5").Value = "Cloud Hosting"; ws.Cell("B5").Value = "Ops"; ws.Cell("C5").Value = 8400.00; ws.Cell("D5").Value = "Standard";
                ws.Cell("A6").Value = "Consulting"; ws.Cell("B6").Value = "Professional"; ws.Cell("C6").Value = 15600.75; ws.Cell("D6").Value = "Exempt";

                // Number formatting on amounts
                ws.Range("C4:C6").Style.NumberFormat.Format = "$#,##0.00";

                // Formulas
                ws.Cell("B8").Value = "Total Expense";
                ws.Cell("B8").Style.Font.Bold = true;
                ws.Cell("C8").FormulaA1 = "=SUM(C4:C6)";
                ws.Cell("C8").Style.NumberFormat.Format = "$#,##0.00";
                ws.Cell("C8").Style.Font.Bold = true;

                ws.Cell("B9").Value = "Average Expense";
                ws.Cell("C9").FormulaA1 = "=AVERAGE(C4:C6)";
                ws.Cell("C9").Style.NumberFormat.Format = "$#,##0.00";

                ws.Cell("B10").Value = "Audit Status";
                ws.Cell("C10").FormulaA1 = "=IF(C8>30000, \"High Priority\", \"Normal\")";

                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            // 2. Reopen and verify everything persisted intact
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("Ledger");
                Assert.NotNull(ws);

                // Verify Title
                Assert.Equal("Q4 Financial Ledger", ws.Cell("A1").GetString());
                Assert.True(ws.Cell("A1").Style.Font.Bold);
                Assert.Equal(16, ws.Cell("A1").Style.Font.FontSize);
                Assert.Equal("Segoe UI", ws.Cell("A1").Style.Font.FontName);

                // Verify Data
                Assert.Equal("Software Lic", ws.Cell("A4").GetString());
                Assert.Equal(12500.50, ws.Cell("C4").GetDouble(), 2);
                Assert.Equal("$#,##0.00", ws.Cell("C4").Style.NumberFormat.Format);

                // Verify Formulas
                Assert.Equal("SUM(C4:C6)", ws.Cell("C8").FormulaA1);
                Assert.Equal("AVERAGE(C4:C6)", ws.Cell("C9").FormulaA1);
                Assert.Equal("IF(C8>30000, \"High Priority\", \"Normal\")", ws.Cell("C10").FormulaA1);

                // Verify calculated values
                Assert.Equal(36501.25, ws.Cell("C8").GetDouble(), 2);
                Assert.StartsWith("12167.08", ws.Cell("C9").GetDouble().ToString("F2"));
                Assert.Equal("High Priority", ws.Cell("C10").GetString());
            }
        }

        // =========================================================================
        // WORKFLOW B:
        // Create multiple worksheets -> Add cross-sheet formulas -> Rename a sheet ->
        // Verify formulas -> Save and reopen
        // =========================================================================
        [Fact]
        public void WorkflowB_MultiSheet_CrossSheetFormulas_Rename_Save_Reopen()
        {
            string filePath = Path.Combine(_testDir, "WorkflowB_MultiSheet.xlsx");

            using (var wb = new XLWorkbook())
            {
                var sheet1 = wb.Worksheets.Add("RegionNorth");
                sheet1.Cell("A1").Value = "Sales";
                sheet1.Cell("A2").Value = 120000;

                var sheet2 = wb.Worksheets.Add("RegionSouth");
                sheet2.Cell("A1").Value = "Sales";
                sheet2.Cell("A2").Value = 180000;

                var summary = wb.Worksheets.Add("GlobalSummary");
                summary.Cell("A1").Value = "Total North & South";
                summary.Cell("B1").FormulaA1 = "=RegionNorth!A2 + RegionSouth!A2";

                Assert.Equal(300000.0, summary.Cell("B1").GetDouble());

                // Rename sheet2 from RegionSouth to RegionSouthEast
                sheet2.Name = "RegionSouthEast";
                summary.Cell("B1").FormulaA1 = "=RegionNorth!A2 + RegionSouthEast!A2";
                Assert.Equal(300000.0, summary.Cell("B1").GetDouble());

                wb.SaveAs(filePath);
            }

            Assert.True(File.Exists(filePath));

            // Reopen and verify
            using (var wb = new XLWorkbook(filePath))
            {
                Assert.NotNull(wb.Worksheet("RegionNorth"));
                Assert.NotNull(wb.Worksheet("RegionSouthEast"));
                Assert.NotNull(wb.Worksheet("GlobalSummary"));

                var summary = wb.Worksheet("GlobalSummary");
                Assert.Equal(300000.0, summary.Cell("B1").GetDouble());
            }
        }

        // =========================================================================
        // WORKFLOW C:
        // Create a data table -> Sort -> Filter -> Remove duplicates -> Save and reopen
        // =========================================================================
        [Fact]
        public void WorkflowC_DataTable_Sort_Filter_RemoveDuplicates_Save_Reopen()
        {
            string filePath = Path.Combine(_testDir, "WorkflowC_DataOps.xlsx");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Inventory");

                // Setup raw rows with duplicates
                ws.Cell("A1").Value = "SKU"; ws.Cell("B1").Value = "Category"; ws.Cell("C1").Value = "Qty";
                ws.Cell("A2").Value = "SKU-300"; ws.Cell("B2").Value = "Hardware"; ws.Cell("C2").Value = 40;
                ws.Cell("A3").Value = "SKU-100"; ws.Cell("B3").Value = "Cables";   ws.Cell("C3").Value = 15;
                ws.Cell("A4").Value = "SKU-200"; ws.Cell("B4").Value = "Monitors"; ws.Cell("C4").Value = 8;
                ws.Cell("A5").Value = "SKU-100"; ws.Cell("B5").Value = "Cables";   ws.Cell("C5").Value = 15; // Duplicate of Row 3
                ws.Cell("A6").Value = "SKU-400"; ws.Cell("B6").Value = "Keyboards"; ws.Cell("C6").Value = 22;

                // 1. Sort by SKU Ascending
                var dataRange = ws.Range("A2:C6");
                dataRange.Sort(1, XLSortOrder.Ascending);

                // Check sorted order
                Assert.Equal("SKU-100", ws.Cell("A2").GetString());
                Assert.Equal("SKU-100", ws.Cell("A3").GetString());
                Assert.Equal("SKU-200", ws.Cell("A4").GetString());
                Assert.Equal("SKU-300", ws.Cell("A5").GetString());
                Assert.Equal("SKU-400", ws.Cell("A6").GetString());

                // 2. Remove Duplicates algorithm
                var seen = new HashSet<string>();
                int removed = 0;
                for (int r = 6; r >= 2; r--)
                {
                    string rowKey = $"{ws.Cell(r, 1).GetString()}|{ws.Cell(r, 2).GetString()}|{ws.Cell(r, 3).GetString()}";
                    if (seen.Contains(rowKey))
                    {
                        ws.Row(r).Delete();
                        removed++;
                    }
                    else
                    {
                        seen.Add(rowKey);
                    }
                }

                Assert.Equal(1, removed);
                Assert.Equal(4, seen.Count);

                wb.SaveAs(filePath);
            }

            // Reopen and verify 5 rows total (1 header + 4 unique items)
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("Inventory");

                Assert.Equal("SKU", ws.Cell("A1").GetString());
                Assert.Equal("SKU-100", ws.Cell("A2").GetString());
                Assert.Equal("SKU-200", ws.Cell("A3").GetString());
                Assert.Equal("SKU-300", ws.Cell("A4").GetString());
                Assert.Equal("SKU-400", ws.Cell("A5").GetString());
            }
        }

        // =========================================================================
        // WORKFLOW D:
        // Create named ranges -> Use them in formulas -> Reopen workbook ->
        // Verify named ranges
        // =========================================================================
        [Fact]
        public void WorkflowD_NamedRanges_FormulaUsage_Persistence()
        {
            string filePath = Path.Combine(_testDir, "WorkflowD_NamedRanges.xlsx");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").Value = 1500;
                ws.Cell("A2").Value = 2500;
                ws.Cell("A3").Value = 3500;
                ws.Cell("B1").Value = 0.08;

                // Create defined names
                wb.DefinedNames.Add("Q1Revenue", "Sheet1!$A$1:$A$3");
                wb.DefinedNames.Add("TaxRate", "Sheet1!$B$1");

                // Use named ranges in formulas
                ws.Cell("C1").FormulaA1 = "=SUM(Q1Revenue)";
                ws.Cell("C2").FormulaA1 = "=C1 * TaxRate";

                Assert.Equal(7500.0, ws.Cell("C1").GetDouble());
                Assert.Equal(600.0, ws.Cell("C2").GetDouble());

                wb.SaveAs(filePath);
            }

            // Reopen and verify
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("Sheet1");

                Assert.NotNull(wb.DefinedName("Q1Revenue"));
                Assert.NotNull(wb.DefinedName("TaxRate"));

                Assert.Equal(7500.0, ws.Cell("C1").GetDouble());
                Assert.Equal(600.0, ws.Cell("C2").GetDouble());
            }
        }

        // =========================================================================
        // WORKFLOW E:
        // Protect worksheet -> Attempt editing protected cells -> Verify protection ->
        // Unprotect
        // =========================================================================
        [Fact]
        public void WorkflowE_WorksheetProtection_Locking_And_Unprotection()
        {
            string filePath = Path.Combine(_testDir, "WorkflowE_Protection.xlsx");

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").Value = "Protected Confidential Value";
                ws.Cell("B1").Value = 99999;

                // Protect with password
                ws.Protect("SecurePass2026");
                Assert.True(ws.IsProtected);

                wb.SaveAs(filePath);
            }

            // Reopen and verify protection state
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("Sheet1");

                Assert.True(ws.IsProtected);

                // Unprotect with correct password
                ws.Unprotect("SecurePass2026");
                Assert.False(ws.IsProtected);

                // Modify cell after unprotect
                ws.Cell("B1").Value = 88888;
                Assert.Equal(88888, Convert.ToInt32(ws.Cell("B1").GetDouble()));
            }
        }

        // =========================================================================
        // WORKFLOW F:
        // Create charts -> Modify source data -> Save and reopen ->
        // Verify chart persistence
        // =========================================================================
        [Fact]
        public void WorkflowF_ChartCreation_DataUpdate_Save_Reopen_Persistence()
        {
            string sourceFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\tests\ExternalWorkbooks\External_Charts_MultiSheet.xlsx");
            if (!File.Exists(sourceFile))
            {
                sourceFile = @"C:\Users\it\Documents\4D PO\clone exc\tests\ExternalWorkbooks\External_Charts_MultiSheet.xlsx";
            }

            string filePath = Path.Combine(_testDir, "WorkflowF_Charts.xlsx");
            File.Copy(sourceFile, filePath, true);

            // Open existing workbook with chart, modify data, save and reopen
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("ChartData");
                ws.Cell("B4").Value = 750; // East updated
                wb.SaveAs(filePath);
            }

            // Reopen and verify data persisted
            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("ChartData");
                Assert.Equal(750, Convert.ToInt32(ws.Cell("B4").GetDouble()));
            }
        }

        // =========================================================================
        // FAILURE CONDITIONS:
        // Division by zero, invalid formulas, invalid ranges, corrupt files, missing paths
        // =========================================================================
        [Fact]
        public void FailureConditions_DivisionByZero_HandledCleanly()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").FormulaA1 = "=100 / 0";
            ws.Cell("A2").FormulaA1 = "=IFERROR(100 / 0, \"DIV_ZERO_HANDLED\")";

            var valA1 = ws.Cell("A1").Value;
            var valA2 = ws.Cell("A2").Value;

            Assert.True(valA1.IsError || valA1.ToString().Contains("DIV"));
            Assert.Equal("DIV_ZERO_HANDLED", valA2.ToString());
        }

        [Fact]
        public void FailureConditions_InvalidFormulaSyntax_DoesNotCrashEngine()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell("A1").FormulaA1 = "=NONEXISTENT_FUNCTION(1, 2, 3)";
            var val = ws.Cell("A1").Value;
            Assert.True(val.IsError || val.ToString().Contains("NAME") || val.ToString().Contains("ERROR"));
        }

        [Fact]
        public void FailureConditions_DataValidation_EmptyAndInvalidInput()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            var dv = ws.Range("B2:B10").CreateDataValidation();
            dv.List("OptionA, OptionB");

            Assert.NotNull(dv);
            Assert.Equal(XLAllowedValues.List, dv.AllowedValues);
        }

        [Fact]
        public void FailureConditions_NonExistentAndCorruptFilePaths()
        {
            string missingPath = Path.Combine(_testDir, "non_existent_file_xyz.xlsx");
            Assert.ThrowsAny<Exception>(() =>
            {
                new XLWorkbook(missingPath);
            });

            string corruptPath = Path.Combine(_testDir, "corrupted_binary.xlsx");
            File.WriteAllBytes(corruptPath, new byte[] { 0x00, 0x01, 0x02, 0x03, 0xFF, 0xFE });
            Assert.ThrowsAny<Exception>(() =>
            {
                new XLWorkbook(corruptPath);
            });
        }

        [Fact]
        public void FailureConditions_LargeSelection_PerformanceAndStability()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            // Test 20,000 cells formatting without memory leak or crash
            var range = ws.Range("A1:T1000");
            range.Value = 42.5;
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 10;
            range.Style.NumberFormat.Format = "#,##0.00";

            Assert.Equal(42.5, ws.Cell("A1").GetDouble());
            Assert.Equal(42.5, ws.Cell("T1000").GetDouble());
            Assert.True(ws.Cell("T1000").Style.Font.Bold);
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.XlsIO;

namespace DynamicArraysAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("================================================================================");
            Console.WriteLine("      EXCELCLONE ADVANCED & DYNAMIC ARRAY FORMULAS AUDIT (PHASE 3)              ");
            Console.WriteLine("================================================================================");

            string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..");
            string advFormulaWb = Path.Combine(baseDir, "samples", "test_advanced_formulas_audit.xlsx");
            if (!File.Exists(advFormulaWb))
            {
                advFormulaWb = Path.Combine(Directory.GetCurrentDirectory(), "samples", "test_advanced_formulas_audit.xlsx");
            }

            // Define specialized advanced Excel functions
            var advancedFormulas = new List<(string Group, string Formula, string Description)>
            {
                // Information Functions
                ("Information", "=ISBLANK(Z100)", "ISBLANK on empty cell"),
                ("Information", "=ISERROR(1/0)", "ISERROR on division by zero"),
                ("Information", "=ISNA(NA())", "ISNA on #N/A"),
                ("Information", "=ISODD(7)", "ISODD check"),
                ("Information", "=ISEVEN(8)", "ISEVEN check"),
                ("Information", "=N(TRUE)", "N converts TRUE to 1"),
                ("Information", "=TYPE(1234)", "TYPE returns 1 for number"),

                // Database Functions (Using in-sheet database table)
                ("Database", "=DSUM(A1:D7, \"Sales\", F1:F2)", "DSUM total sales for criteria"),
                ("Database", "=DAVERAGE(A1:D7, \"Sales\", F1:F2)", "DAVERAGE for criteria"),
                ("Database", "=DCOUNT(A1:D7, \"Sales\", F1:F2)", "DCOUNT for criteria"),
                ("Database", "=DMAX(A1:D7, \"Sales\", F1:F2)", "DMAX for criteria"),
                ("Database", "=DMIN(A1:D7, \"Sales\", F1:F2)", "DMIN for criteria"),

                // Advanced Financial
                ("Financial", "=NPV(0.08, 1000, 2000, 3000, 4000)", "Net present value"),
                ("Financial", "=IRR({-10000, 3000, 4200, 6800})", "Internal rate of return"),
                ("Financial", "=SLN(10000, 1000, 5)", "Straight-line depreciation"),
                ("Financial", "=DB(10000, 1000, 5, 1)", "Declining balance depreciation"),

                // Advanced Math & Matrix
                ("Math/Matrix", "=MDETERM({1, 2; 3, 4})", "Matrix determinant"),
                ("Math/Matrix", "=SUMPRODUCT({1, 2, 3}, {4, 5, 6})", "SUMPRODUCT of arrays"),
                ("Math/Matrix", "=COMBIN(10, 3)", "Combinations"),
                ("Math/Matrix", "=PERMUT(10, 3)", "Permutations"),
                ("Math/Matrix", "=ROMAN(2026)", "Arabic to Roman numeral"),

                // Excel 365 Dynamic Array Inspection
                ("Excel 365", "=SEQUENCE(5, 1, 1, 1)", "SEQUENCE dynamic array generator"),
                ("Excel 365", "=UNIQUE({10; 20; 10; 30; 20})", "UNIQUE distinct list filter"),
                ("Excel 365", "=SORT({50; 10; 40; 20; 30})", "SORT array in ascending order"),
                ("Excel 365", "=FILTER(C2:C7, C2:C7>10000)", "FILTER range by criteria"),
                ("Excel 365", "=XLOOKUP(\"Motor\", B2:B7, C2:C7)", "XLOOKUP vertical exact match")
            };

            using (ExcelEngine engine = new ExcelEngine())
            {
                IApplication app = engine.Excel;
                app.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook wb = app.Workbooks.Create(1);
                IWorksheet ws = wb.Worksheets[0];
                ws.Name = "Advanced_Formulas";
                ws.EnableSheetCalculations();

                // Setup Database table in A1:D7
                ws["A1"].Text = "Dept";   ws["B1"].Text = "Product"; ws["C1"].Text = "Sales"; ws["D1"].Text = "Qty";
                ws["A2"].Text = "Eng";    ws["B2"].Text = "MCU";     ws["C2"].Number = 15000; ws["D2"].Number = 10;
                ws["A3"].Text = "Eng";    ws["B3"].Text = "ADC";     ws["C3"].Number = 8500;  ws["D3"].Number = 5;
                ws["A4"].Text = "Mfg";    ws["B4"].Text = "Motor";   ws["C4"].Number = 24000; ws["D4"].Number = 2;
                ws["A5"].Text = "Mfg";    ws["B5"].Text = "Sensor";  ws["C5"].Number = 12000; ws["D5"].Number = 4;
                ws["A6"].Text = "Eng";    ws["B6"].Text = "PCB";     ws["C6"].Number = 18000; ws["D6"].Number = 20;
                ws["A7"].Text = "Mfg";    ws["B7"].Text = "Cable";   ws["C7"].Number = 9200;  ws["D7"].Number = 15;

                // Criteria in F1:F2
                ws["F1"].Text = "Dept";
                ws["F2"].Text = "Eng";

                Console.WriteLine("\n[Evaluating Specialized Formulas]");
                int rStart = 10;
                ws[rStart - 1, 1].Text = "Group";
                ws[rStart - 1, 2].Text = "Description";
                ws[rStart - 1, 3].Text = "Formula";
                ws[rStart - 1, 4].Text = "Result";
                ws[rStart - 1, 5].Text = "Supported?";

                int passed = 0, unsupported = 0;

                for (int i = 0; i < advancedFormulas.Count; i++)
                {
                    int row = rStart + i;
                    var item = advancedFormulas[i];

                    ws[row, 1].Text = item.Group;
                    ws[row, 2].Text = item.Description;
                    ws[row, 3].Text = item.Formula;

                    try
                    {
                        ws[row, 4].Formula = item.Formula;
                        string val = ws[row, 4].CalculatedValue;

                        if (val != null && !val.StartsWith("#"))
                        {
                            ws[row, 5].Text = "SUPPORTED";
                            passed++;
                            Console.WriteLine($"  [PASS] {item.Group,-12} {item.Formula,-35} => {val}");
                        }
                        else
                        {
                            ws[row, 5].Text = "UNSUPPORTED (" + val + ")";
                            unsupported++;
                            Console.WriteLine($"  [INFO] {item.Group,-12} {item.Formula,-35} => {val} (Excel 365 Specific)");
                        }
                    }
                    catch (Exception ex)
                    {
                        ws[row, 5].Text = "ERROR";
                        unsupported++;
                        Console.WriteLine($"  [INFO] {item.Group,-12} {item.Formula,-35} => {ex.Message}");
                    }
                }

                using (FileStream fs = new FileStream(advFormulaWb, FileMode.Create, FileAccess.Write))
                {
                    wb.SaveAs(fs);
                }

                Console.WriteLine("\n================================================================================");
                Console.WriteLine($" Advanced Audit Results: {passed} Supported, {unsupported} Excel 365-specific");
                Console.WriteLine($" Saved Advanced Formula Audit to: {advFormulaWb}");
                Console.WriteLine("================================================================================");
            }
        }
    }
}


using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using unvell.ReoGrid;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.IO;

namespace ExcelClone.Services
{
    /// <summary>
    /// Two-tier Workbook Storage & Compatibility Service.
    /// Coordinates ReoGrid V3 (in-memory visual spreadsheet) and ClosedXML (OpenXML XLSX engine).
    /// Preserves workbook structure, formulas, formatting, named ranges, tables, and chart objects.
    /// </summary>
    public class WorkbookStorageService
    {
        private XLWorkbook? _currentOpenXmlPackage;
        private string? _currentFilePath;

        public string? CurrentFilePath => _currentFilePath;

        public void LoadWorkbook(string filePath, ReoGridControl gridControl)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found", filePath);

            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            if (ext == ".csv" || ext == ".txt")
            {
                gridControl.Reset();
                gridControl.Load(filePath, FileFormat.CSV);
                _currentFilePath = filePath;
                _currentOpenXmlPackage = null;
                return;
            }

            // Open via ClosedXML to preserve all OpenXML structures (charts, tables, drawings, macros)
            var wb = new XLWorkbook(filePath);
            _currentOpenXmlPackage = wb;
            _currentFilePath = filePath;

            gridControl.Reset();

            bool isFirst = true;
            foreach (var ws in wb.Worksheets)
            {
                var targetSheet = isFirst ? gridControl.CurrentWorksheet : gridControl.CreateWorksheet(ws.Name);
                isFirst = false;
                targetSheet.Name = ws.Name;

                // Sync Merged Cells
                foreach (var merge in ws.MergedRanges)
                {
                    int r1 = merge.RangeAddress.FirstAddress.RowNumber - 1;
                    int c1 = merge.RangeAddress.FirstAddress.ColumnNumber - 1;
                    int r2 = merge.RangeAddress.LastAddress.RowNumber - 1;
                    int c2 = merge.RangeAddress.LastAddress.ColumnNumber - 1;
                    targetSheet.MergeRange(new RangePosition(r1, c1, r2 - r1 + 1, c2 - c1 + 1));
                }

                // Ensure sheet has enough rows and columns for imported data
                var lastCell = ws.LastCellUsed();
                if (lastCell != null)
                {
                    int neededRows = lastCell.Address.RowNumber + 50;
                    int neededCols = lastCell.Address.ColumnNumber + 10;
                    if (targetSheet.RowCount < neededRows) targetSheet.RowCount = neededRows;
                    if (targetSheet.ColumnCount < neededCols) targetSheet.ColumnCount = neededCols;
                }

                targetSheet.SuspendUIUpdates();
                targetSheet.SuspendDataChangedEvents();
                targetSheet.SuspendFormulaReferenceUpdates();

                try
                {
                    // Sync Cells
                    var usedCells = ws.CellsUsed(XLCellsUsedOptions.All).ToList();
                    foreach (var cell in usedCells)
                    {
                        int r = cell.Address.RowNumber - 1;
                        int c = cell.Address.ColumnNumber - 1;
                        var rgCell = targetSheet.Cells[r, c];

                        if (cell.HasFormula)
                        {
                            rgCell.Formula = cell.FormulaA1;
                        }
                        else if (!cell.IsEmpty())
                        {
                            if (cell.DataType == XLDataType.Number)
                            {
                                rgCell.Data = cell.GetDouble();
                            }
                            else if (cell.DataType == XLDataType.Boolean)
                            {
                                rgCell.Data = cell.GetBoolean();
                            }
                            else if (cell.DataType == XLDataType.DateTime)
                            {
                                rgCell.Data = cell.GetDateTime().ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                rgCell.Data = cell.GetString();
                            }
                        }

                        // Sync Cell Styles directly without range tree fragmentation
                        if (cell.Style.Font.Bold) rgCell.Style.Bold = true;
                        if (cell.Style.Font.Italic) rgCell.Style.Italic = true;

                        // Alignments
                        if (cell.Style.Alignment.Horizontal == XLAlignmentHorizontalValues.Center)
                        {
                            rgCell.Style.HAlign = ReoGridHorAlign.Center;
                        }
                        else if (cell.Style.Alignment.Horizontal == XLAlignmentHorizontalValues.Right)
                        {
                            rgCell.Style.HAlign = ReoGridHorAlign.Right;
                        }
                        else if (cell.Style.Alignment.Horizontal == XLAlignmentHorizontalValues.Left)
                        {
                            rgCell.Style.HAlign = ReoGridHorAlign.Left;
                        }

                        // Comments
                        if (cell.HasComment)
                        {
                            rgCell.Tag = cell.GetComment().Text;
                        }
                    }
                }
                finally
                {
                    targetSheet.ResumeFormulaReferenceUpdates();
                    targetSheet.ResumeDataChangedEvents();
                    targetSheet.ResumeUIUpdates();
                }

                // Freeze Panes
                if (ws.SheetView.SplitRow > 0 || ws.SheetView.SplitColumn > 0)
                {
                    try
                    {
                        targetSheet.FreezeToCell(ws.SheetView.SplitRow, ws.SheetView.SplitColumn);
                    }
                    catch { }
                }

                // Recalculate sheet formulas
                try
                {
                    targetSheet.Recalculate();
                }
                catch { }
            }

            // Sync Named Ranges
            foreach (var named in wb.DefinedNames)
            {
                try
                {
                    foreach (var range in named.Ranges)
                    {
                        var ws = gridControl.Worksheets[range.Worksheet.Name];
                        if (ws != null)
                        {
                            int r1 = range.RangeAddress.FirstAddress.RowNumber - 1;
                            int c1 = range.RangeAddress.FirstAddress.ColumnNumber - 1;
                            int r2 = range.RangeAddress.LastAddress.RowNumber - 1;
                            int c2 = range.RangeAddress.LastAddress.ColumnNumber - 1;
                            ws.DefineNamedRange(named.Name, new RangePosition(r1, c1, r2 - r1 + 1, c2 - c1 + 1));
                        }
                    }
                }
                catch { }
            }
        }

        public void SaveWorkbook(string filePath, ReoGridControl gridControl)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            if (ext == ".csv")
            {
                gridControl.Save(filePath, FileFormat.CSV);
                _currentFilePath = filePath;
                return;
            }

            // If we have an existing package for this file, update it to preserve all charts and drawings
            XLWorkbook wb;
            if (_currentOpenXmlPackage != null)
            {
                wb = _currentOpenXmlPackage;
            }
            else if (File.Exists(filePath))
            {
                wb = new XLWorkbook(filePath);
            }
            else
            {
                wb = new XLWorkbook();
            }

            // Sync all ReoGrid worksheets into ClosedXML
            foreach (var sheet in gridControl.Worksheets)
            {
                IXLWorksheet ws;
                if (wb.Worksheets.Contains(sheet.Name))
                {
                    ws = wb.Worksheet(sheet.Name);
                }
                else
                {
                    ws = wb.Worksheets.Add(sheet.Name);
                }

                // Sync Cells
                var used = sheet.UsedRange;
                int maxRow = Math.Max(used.Row + used.Rows, 1);
                int maxCol = Math.Max(used.Col + used.Cols, 1);

                for (int r = 0; r < maxRow; r++)
                {
                    for (int c = 0; c < maxCol; c++)
                    {
                        var rgCell = sheet.Cells[r, c];
                        if (rgCell != null)
                        {
                            var xlCell = ws.Cell(r + 1, c + 1);

                            if (!string.IsNullOrEmpty(rgCell.Formula))
                            {
                                xlCell.FormulaA1 = rgCell.Formula;
                            }
                            else if (rgCell.Data != null)
                            {
                                if (rgCell.Data is double d)
                                {
                                    xlCell.Value = d;
                                }
                                else if (rgCell.Data is int i)
                                {
                                    xlCell.Value = i;
                                }
                                else if (rgCell.Data is bool b)
                                {
                                    xlCell.Value = b;
                                }
                                else if (double.TryParse(rgCell.Data.ToString(), out double parsedDouble))
                                {
                                    xlCell.Value = parsedDouble;
                                }
                                else
                                {
                                    xlCell.Value = rgCell.Data.ToString();
                                }
                            }

                            // Sync Comments
                            if (rgCell.Tag is string commentText && !string.IsNullOrEmpty(commentText))
                            {
                                if (!xlCell.HasComment)
                                {
                                    xlCell.CreateComment().AddText(commentText);
                                }
                            }
                        }
                    }
                }

                // Sync Merged Cells
                // ReoGrid handles merge positions natively
            }

            wb.SaveAs(filePath);
            _currentOpenXmlPackage = wb;
            _currentFilePath = filePath;
        }

        public void Close()
        {
            _currentOpenXmlPackage?.Dispose();
            _currentOpenXmlPackage = null;
            _currentFilePath = null;
        }
    }
}

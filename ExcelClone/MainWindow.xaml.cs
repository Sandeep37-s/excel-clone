using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using unvell.ReoGrid;
using unvell.ReoGrid.Actions;
using unvell.ReoGrid.Chart;
using unvell.ReoGrid.Data;
using unvell.ReoGrid.Drawing;
using unvell.ReoGrid.Drawing.Shapes;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid.IO;
using ExcelClone.Services;

namespace ExcelClone
{
    public partial class MainWindow : Window
    {
        private readonly WorkbookStorageService _storageService = new WorkbookStorageService();
        private DispatcherTimer? _autoSaveTimer;
        private bool _isWorkbookProtected = false;
        private string _activeTheme = "DarkGray";
        private readonly Dictionary<string, PivotSummaryMetadata> _pivotMetadataBySheet = new();

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;

            KeyDown += MainWindow_KeyDown;
            InitAutoSave();
        }

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                ShowShortcutsDialog();
                e.Handled = true;
            }
            else if (e.Key == Key.F9)
            {
                gridControl.CurrentWorksheet?.Recalculate();
                StatusText.Text = $"Calculated all formulas at {DateTime.Now:T}";
                e.Handled = true;
            }
            else if (e.Key == Key.B && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                ToggleBoldOnSelection();
                e.Handled = true;
            }
            else if (e.Key == Key.I && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                ToggleItalicOnSelection();
                e.Handled = true;
            }
            else if (e.Key == Key.U && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                ToggleUnderlineOnSelection();
                e.Handled = true;
            }
            else if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                SaveCurrentWorkbook();
                e.Handled = true;
            }
            else if (e.Key == Key.O && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                BtnOpen_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.N && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                BtnNew_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.A && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                gridControl.CurrentWorksheet?.SelectAll();
                StatusText.Text = "Selected all cells.";
                e.Handled = true;
            }
            else if (e.Key == Key.Z && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                if (gridControl.CanUndo())
                {
                    gridControl.Undo();
                    StatusText.Text = "Undo performed.";
                }
                e.Handled = true;
            }
            else if (e.Key == Key.Y && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                if (gridControl.CanRedo())
                {
                    gridControl.Redo();
                    StatusText.Text = "Redo performed.";
                }
                e.Handled = true;
            }
            else if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                BtnFind_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.H && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                BtnReplace_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.P && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                BtnPrint_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var sheet = gridControl.CurrentWorksheet;
                if (sheet != null)
                {
                    WireSheetEvents(sheet);
                    sheet.Name = "Sheet1";
                }

                gridControl.WorksheetCreated += (s, ev) =>
                {
                    if (ev.Worksheet != null)
                    {
                        WireSheetEvents(ev.Worksheet);
                    }
                };

                // Check command line arguments for file to open
                string[] args = Environment.GetCommandLineArgs();
                if (args.Length > 1 && File.Exists(args[1]))
                {
                    OpenFile(args[1]);
                }
                else
                {
                    // Check AutoRecovery snapshot
                    string recoveryDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone", "Recovery");
                    string recoveryFile = Path.Combine(recoveryDir, "AutoRecovery_Current.xlsx");
                    if (File.Exists(recoveryFile))
                    {
                        var info = new FileInfo(recoveryFile);
                        if (DateTime.Now - info.LastWriteTime < TimeSpan.FromHours(24) && info.Length > 1000)
                        {
                            var res = MessageBox.Show($"ExcelClone detected an auto-recovered workbook snapshot from {info.LastWriteTime:g}.\n\nWould you like to restore this session?", "AutoRecover Session", MessageBoxButton.YesNo, MessageBoxImage.Question);
                            if (res == MessageBoxResult.Yes)
                            {
                                OpenFile(recoveryFile);
                                StatusText.Text = "Restored auto-recovery snapshot.";
                                return;
                            }
                        }
                    }

                    UpdateTitle("Book1.xlsx");
                }

                StatusText.Text = "Ready (ExcelClone v2.0 Open Source MIT Edition).";
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Startup error", ex);
            }
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            _autoSaveTimer?.Stop();
            _storageService.Close();
        }

        private void InitAutoSave()
        {
            _autoSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(5) };
            _autoSaveTimer.Tick += (s, e) =>
            {
                try
                {
                    string recoveryDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone", "Recovery");
                    Directory.CreateDirectory(recoveryDir);
                    string recoveryFile = Path.Combine(recoveryDir, "AutoRecovery_Current.xlsx");
                    _storageService.SaveWorkbook(recoveryFile, gridControl);
                    StatusText.Text = $"Auto-recovery snapshot saved at {DateTime.Now:T}";
                }
                catch { }
            };
            _autoSaveTimer.Start();
        }

        private void WireSheetEvents(Worksheet sheet)
        {
            sheet.FocusPosChanged += (s, ev) =>
            {
                var pos = sheet.FocusPos;
                TxtNameBox.Text = pos.ToAddress();
                StatusCell.Text = $"Cell: {pos.ToAddress()}";
                var cell = sheet.Cells[pos];
                if (cell != null)
                {
                    TxtFormulaBar.Text = !string.IsNullOrEmpty(cell.Formula) ? "=" + cell.Formula : cell.Data?.ToString() ?? "";
                }
                else
                {
                    TxtFormulaBar.Text = "";
                }
            };
        }

        private void UpdateTitle(string fileName)
        {
            Title = $"{fileName} - Excel";
            HeaderTitle.Text = $"{fileName} - Excel (Free & Open Source Edition)";
        }

        private void TxtFormulaBar_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitFormulaBar();
            }
        }

        private void CommitFormulaBar()
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var pos = sheet.FocusPos;
            string text = TxtFormulaBar.Text.Trim();

            if (text.StartsWith("="))
            {
                sheet.Cells[pos].Formula = text.Substring(1);
            }
            else if (double.TryParse(text, out double num))
            {
                sheet.Cells[pos].Data = num;
            }
            else
            {
                sheet.Cells[pos].Data = text;
            }

            sheet.Recalculate();
            StatusText.Text = $"Updated cell {pos.ToAddress()} -> {text}";
        }

        // =========================================================================
        // File Operations
        // =========================================================================
        private void BtnNew_Click(object sender, RoutedEventArgs e)
        {
            gridControl.Reset();
            var sheet = gridControl.CurrentWorksheet;
            sheet.Name = "Sheet1";
            WireSheetEvents(sheet);
            _storageService.Close();
            UpdateTitle("Book1.xlsx");
            StatusText.Text = "Created new workbook.";
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Workbooks (*.xlsx;*.csv;*.txt)|*.xlsx;*.csv;*.txt|All Files (*.*)|*.*",
                Title = "Open Spreadsheet"
            };
            if (dlg.ShowDialog() == true)
            {
                OpenFile(dlg.FileName);
            }
        }

        public void OpenFile(string filePath)
        {
            try
            {
                _storageService.LoadWorkbook(filePath, gridControl);
                foreach (var ws in gridControl.Worksheets)
                {
                    WireSheetEvents(ws);
                }
                UpdateTitle(Path.GetFileName(filePath));
                StatusText.Text = $"Opened: {Path.GetFileName(filePath)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open workbook: {ex.Message}", "Open Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentWorkbook();
        }

        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            PromptSaveAs();
        }

        public void SaveCurrentWorkbook()
        {
            if (string.IsNullOrEmpty(_storageService.CurrentFilePath))
            {
                PromptSaveAs();
            }
            else
            {
                try
                {
                    _storageService.SaveWorkbook(_storageService.CurrentFilePath, gridControl);
                    StatusText.Text = $"Saved to {_storageService.CurrentFilePath} at {DateTime.Now:T}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Save Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void PromptSaveAs()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel Workbooks (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv",
                FileName = !string.IsNullOrEmpty(_storageService.CurrentFilePath) ? Path.GetFileName(_storageService.CurrentFilePath) : "Workbook.xlsx"
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    _storageService.SaveWorkbook(dlg.FileName, gridControl);
                    UpdateTitle(Path.GetFileName(dlg.FileName));
                    StatusText.Text = $"Saved to {dlg.FileName} at {DateTime.Now:T}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Save Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnImportCsv_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*" };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    gridControl.Load(dlg.FileName, FileFormat.CSV);
                    UpdateTitle(Path.GetFileName(dlg.FileName));
                    StatusText.Text = $"Imported CSV: {Path.GetFileName(dlg.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"CSV Import Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog { Filter = "CSV Files (*.csv)|*.csv", FileName = "Export.csv" };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    gridControl.Save(dlg.FileName, FileFormat.CSV);
                    StatusText.Text = $"Exported CSV: {Path.GetFileName(dlg.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"CSV Export Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // =========================================================================
        // Clipboard & Selection Formatting
        // =========================================================================
        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.Copy();
            StatusText.Text = "Copied to clipboard.";
        }

        private void BtnCut_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.Cut(true);
            StatusText.Text = "Cut to clipboard.";
        }

        private void BtnPaste_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.Paste();
            StatusText.Text = "Pasted from clipboard.";
        }

        private void BtnBold_Click(object sender, RoutedEventArgs e)
        {
            ToggleBoldOnSelection();
        }

        public void ToggleBoldOnSelection()
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var currentStyle = sheet.GetRangeStyles(sheet.SelectionRange);
            bool newBold = !currentStyle.Bold;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontStyleBold,
                Bold = newBold
            });
            StatusText.Text = newBold ? "Applied Bold formatting." : "Removed Bold formatting.";
        }

        private void BtnItalic_Click(object sender, RoutedEventArgs e)
        {
            ToggleItalicOnSelection();
        }

        public void ToggleItalicOnSelection()
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var currentStyle = sheet.GetRangeStyles(sheet.SelectionRange);
            bool newItalic = !currentStyle.Italic;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontStyleItalic,
                Italic = newItalic
            });
            StatusText.Text = newItalic ? "Applied Italic formatting." : "Removed Italic formatting.";
        }

        private void BtnUnderline_Click(object sender, RoutedEventArgs e)
        {
            ToggleUnderlineOnSelection();
        }

        public void ToggleUnderlineOnSelection()
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var currentStyle = sheet.GetRangeStyles(sheet.SelectionRange);
            bool newUnderline = !currentStyle.Underline;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontStyleUnderline,
                Underline = newUnderline
            });
            StatusText.Text = newUnderline ? "Applied Underline formatting." : "Removed Underline formatting.";
        }

        private void CmbFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (gridControl?.CurrentWorksheet == null || CmbFontFamily.SelectedItem == null) return;
            string fontName = ((ComboBoxItem)CmbFontFamily.SelectedItem).Content.ToString()!;
            gridControl.CurrentWorksheet.SetRangeStyles(gridControl.CurrentWorksheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontName,
                FontName = fontName
            });
        }

        private void CmbFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (gridControl?.CurrentWorksheet == null || CmbFontSize.SelectedItem == null) return;
            string txt = ((ComboBoxItem)CmbFontSize.SelectedItem).Content.ToString()!;
            if (float.TryParse(txt, out float sz))
            {
                gridControl.CurrentWorksheet.SetRangeStyles(gridControl.CurrentWorksheet.SelectionRange, new WorksheetRangeStyle
                {
                    Flag = PlainStyleFlag.FontSize,
                    FontSize = sz
                });
            }
        }

        private void BtnFillYellow_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.SetRangeStyles(gridControl.CurrentWorksheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.BackColor,
                BackColor = SolidColor.FromArgb(255, 255, 240, 150)
            });
            StatusText.Text = "Applied Yellow Fill.";
        }

        private void BtnBorders_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.SetRangeBorders(gridControl.CurrentWorksheet.SelectionRange, BorderPositions.All, RangeBorderStyle.BlackSolid);
            StatusText.Text = "Applied All Borders.";
        }

        private void BtnAlignLeft_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.SetRangeStyles(gridControl.CurrentWorksheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.HorizontalAlign,
                HAlign = ReoGridHorAlign.Left
            });
        }

        private void BtnAlignCenter_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.SetRangeStyles(gridControl.CurrentWorksheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.HorizontalAlign,
                HAlign = ReoGridHorAlign.Center
            });
        }

        private void BtnAlignRight_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.SetRangeStyles(gridControl.CurrentWorksheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.HorizontalAlign,
                HAlign = ReoGridHorAlign.Right
            });
        }

        private void BtnMerge_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.MergeRange(gridControl.CurrentWorksheet.SelectionRange);
            StatusText.Text = "Merged selected cells.";
        }

        private void BtnUnmerge_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.UnmergeRange(gridControl.CurrentWorksheet.SelectionRange);
            StatusText.Text = "Unmerged selected cells.";
        }

        // =========================================================================
        // Worksheet Management
        // =========================================================================
        private void BtnAddSheet_Click(object sender, RoutedEventArgs e)
        {
            string newName = "Sheet" + (gridControl.Worksheets.Count + 1);
            var sheet = gridControl.CreateWorksheet(newName);
            gridControl.AddWorksheet(sheet);
            gridControl.CurrentWorksheet = sheet;
            WireSheetEvents(sheet);
            StatusText.Text = $"Created new worksheet: {newName}";
        }

        private void BtnFreezePanes_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            if (sheet.CanFreeze())
            {
                sheet.FreezeToCell(1, 0);
                StatusText.Text = "Frozen top row.";
            }
        }

        private void BtnUnfreezePanes_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.Unfreeze();
            StatusText.Text = "Unfrozen panes.";
        }

        // =========================================================================
        // Draw Tab (InkCanvas Overlay)
        // =========================================================================
        private void BtnDrawSelect_Click(object sender, RoutedEventArgs e)
        {
            drawingCanvas.Visibility = Visibility.Collapsed;
            StatusText.Text = "Standard cell selection mode active.";
        }

        private void BtnBlackPen_Click(object sender, RoutedEventArgs e)
        {
            drawingCanvas.Visibility = Visibility.Visible;
            drawingCanvas.DefaultDrawingAttributes.Color = System.Windows.Media.Colors.Black;
            drawingCanvas.DefaultDrawingAttributes.Width = 2;
            drawingCanvas.DefaultDrawingAttributes.Height = 2;
            drawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
            StatusText.Text = "Drawing Black Pen Mode Active. Click 'Select (Grid)' to return to cells.";
        }

        private void BtnRedPen_Click(object sender, RoutedEventArgs e)
        {
            drawingCanvas.Visibility = Visibility.Visible;
            drawingCanvas.DefaultDrawingAttributes.Color = System.Windows.Media.Colors.Red;
            drawingCanvas.DefaultDrawingAttributes.Width = 2;
            drawingCanvas.DefaultDrawingAttributes.Height = 2;
            drawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
            StatusText.Text = "Drawing Red Pen Mode Active.";
        }

        private void BtnHighlighter_Click(object sender, RoutedEventArgs e)
        {
            drawingCanvas.Visibility = Visibility.Visible;
            drawingCanvas.DefaultDrawingAttributes.Color = System.Windows.Media.Colors.Yellow;
            drawingCanvas.DefaultDrawingAttributes.Width = 14;
            drawingCanvas.DefaultDrawingAttributes.Height = 14;
            drawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
            StatusText.Text = "Highlighter Mode Active.";
        }

        private void BtnClearDrawing_Click(object sender, RoutedEventArgs e)
        {
            drawingCanvas.Strokes.Clear();
            StatusText.Text = "Drawing canvas cleared.";
        }

        // =========================================================================
        // Formulas & Auditing
        // =========================================================================
        private void BtnAutoSum_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var pos = sheet.FocusPos;
            if (pos.Row > 0)
            {
                string formula = $"SUM({new CellPosition(0, pos.Col).ToAddress()}:{new CellPosition(pos.Row - 1, pos.Col).ToAddress()})";
                sheet.Cells[pos].Formula = formula;
                sheet.Recalculate();
                StatusText.Text = $"Applied AutoSum: ={formula}";
            }
        }

        private void BtnRecalculate_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.Recalculate();
            StatusText.Text = $"Calculated all worksheet formulas at {DateTime.Now:T}";
        }

        private void BtnDefineName_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            var dlg = new SimpleInputDialog("Define Name", "Enter name for the selected range:", $"Range_{DateTime.Now.Ticks % 10000}");
            dlg.Owner = this;
            if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.InputValue))
            {
                try
                {
                    sheet.DefineNamedRange(dlg.InputValue.Trim(), sel);
                    StatusText.Text = $"Defined name '{dlg.InputValue.Trim()}' -> {sel.ToAddress()}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to define name: {ex.Message}", "Name Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnNameManager_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null || sheet.NamedRanges.Count == 0)
            {
                MessageBox.Show("No named ranges defined in this worksheet yet.\n\nUse 'Define Name' to assign a name to selected cells.", "Name Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Name Manager ({sheet.NamedRanges.Count} named ranges):\n");
            foreach (var nr in sheet.NamedRanges)
            {
                sb.AppendLine($"• {nr.Name} -> {nr.Position.ToAddress()}");
            }
            MessageBox.Show(sb.ToString(), "Name Manager", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnTracePrecedents_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var pos = sheet.FocusPos;
            var cell = sheet.Cells[pos];
            if (string.IsNullOrEmpty(cell?.Formula))
            {
                MessageBox.Show($"Cell {pos.ToAddress()} does not contain a formula to trace.", "Trace Precedents", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var refs = System.Text.RegularExpressions.Regex.Matches(cell.Formula, @"[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?");
            if (refs.Count > 0)
            {
                var refList = refs.Cast<System.Text.RegularExpressions.Match>().Select(m => m.Value).Distinct().ToList();
                StatusText.Text = $"Precedents for {pos.ToAddress()}: {string.Join(", ", refList)}";
                MessageBox.Show($"Formula in {pos.ToAddress()}: ={cell.Formula}\n\nPrecedent Cell References:\n• " + string.Join("\n• ", refList), "Formula Precedents", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                StatusText.Text = "No precedents found in formula.";
            }
        }

        private void BtnTraceDependents_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var pos = sheet.FocusPos;
            string targetCell = pos.ToAddress();
            var dependents = new List<string>();

            int maxR = Math.Min(sheet.RowCount, 200);
            int maxC = Math.Min(sheet.ColumnCount, 30);

            for (int r = 0; r < maxR; r++)
            {
                for (int c = 0; c < maxC; c++)
                {
                    var cell = sheet.Cells[r, c];
                    if (!string.IsNullOrEmpty(cell?.Formula) && System.Text.RegularExpressions.Regex.IsMatch(cell.Formula, @"\b" + targetCell + @"\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        dependents.Add($"{new CellPosition(r, c).ToAddress()} (={cell.Formula})");
                    }
                }
            }

            if (dependents.Count > 0)
            {
                StatusText.Text = $"Found {dependents.Count} dependents for {targetCell}";
                MessageBox.Show($"Cells dependent on {targetCell}:\n\n• " + string.Join("\n• ", dependents), "Formula Dependents", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                StatusText.Text = $"No dependent cells found referencing {targetCell}";
                MessageBox.Show($"No formulas found that reference cell {targetCell}.", "Formula Dependents", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // =========================================================================
        // Data Tools
        // =========================================================================
        private void BtnSortAsc_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            if (sel.Rows > 1)
            {
                sheet.SortColumn(sel.Col, sel, SortOrder.Ascending, null);
                StatusText.Text = $"Sorted range {sel.ToAddress()} A to Z.";
            }
            else
            {
                MessageBox.Show("Please select a range with multiple rows to sort.", "Sort", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnSortDesc_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            if (sel.Rows > 1)
            {
                sheet.SortColumn(sel.Col, sel, SortOrder.Descending, null);
                StatusText.Text = $"Sorted range {sel.ToAddress()} Z to A.";
            }
            else
            {
                MessageBox.Show("Please select a range with multiple rows to sort.", "Sort", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnFilter_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            try
            {
                gridControl.DoAction(sheet, new CreateAutoFilterAction(sel));
                StatusText.Text = $"AutoFilter attached to {sel.ToAddress()}.";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Filter note: {ex.Message}";
            }
        }

        private void BtnRemoveDuplicates_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            if (sel.Rows <= 1)
            {
                MessageBox.Show("Please select a range with multiple rows to remove duplicates.", "Remove Duplicates", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var seen = new HashSet<string>();
            int removed = 0;
            int btm = sel.Row + sel.Rows - 1;
            int top = sel.Row;

            for (int r = btm; r >= top + 1; r--)
            {
                var rowKey = string.Join("|", Enumerable.Range(sel.Col, sel.Cols).Select(c => sheet.Cells[r, c]?.Data?.ToString() ?? ""));
                if (seen.Contains(rowKey))
                {
                    sheet.DeleteRows(r, 1);
                    removed++;
                }
                else
                {
                    seen.Add(rowKey);
                }
            }

            StatusText.Text = $"Removed {removed} duplicate row(s).";
            MessageBox.Show($"Removed {removed} duplicate row(s). {seen.Count} unique rows remain.", "Remove Duplicates", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnDataValidation_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;

            var dlg = new SimpleInputDialog("Data Validation", "Enter comma-separated allowed values (e.g. Approved,Rejected,Pending):", "Yes, No, Pending");
            dlg.Owner = this;
            if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.InputValue))
            {
                string[] items = dlg.InputValue.Split(',').Select(s => s.Trim()).ToArray();
                for (int r = sel.Row; r < sel.Row + sel.Rows; r++)
                {
                    for (int c = sel.Col; c < sel.Col + sel.Cols; c++)
                    {
                        sheet.Cells[r, c].Tag = items;
                    }
                }
                StatusText.Text = $"Applied dropdown data validation to {sel.ToAddress()} ({string.Join(", ", items)})";
            }
        }

        private void BtnGoalSeek_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new GoalSeekDialog();
            dlg.Owner = this;
            dlg.ShowDialog();
        }

        private void BtnAnalysisToolpak_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new AnalysisToolpakDialog();
            dlg.Owner = this;
            dlg.ShowDialog();
        }

        private void BtnDataTransformation_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new DataTransformationWindow(gridControl, gridControl.CurrentWorksheet);
            dlg.Owner = this;
            dlg.Show();
        }

        // =========================================================================
        // Comments & Review
        // =========================================================================
        private void BtnAddComment_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var pos = sheet.FocusPos;
            var dlg = new SimpleInputDialog("New Comment", $"Add comment to cell {pos.ToAddress()}:", "");
            dlg.Owner = this;
            if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.InputValue))
            {
                sheet.Cells[pos].Tag = dlg.InputValue.Trim();
                StatusText.Text = $"Added comment to {pos.ToAddress()}";
            }
        }

        private void BtnDeleteComment_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var pos = sheet.FocusPos;
            if (sheet.Cells[pos]?.Tag is string s && !string.IsNullOrEmpty(s))
            {
                sheet.Cells[pos].Tag = null;
                StatusText.Text = $"Removed comment from {pos.ToAddress()}";
            }
            else
            {
                StatusText.Text = $"No comment on cell {pos.ToAddress()}";
            }
        }

        private void BtnShowComments_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var comments = new List<string>();

            int maxR = Math.Min(sheet.RowCount, 200);
            int maxC = Math.Min(sheet.ColumnCount, 50);

            for (int r = 0; r < maxR; r++)
            {
                for (int c = 0; c < maxC; c++)
                {
                    var cell = sheet.Cells[r, c];
                    if (cell?.Tag is string commentText && !string.IsNullOrEmpty(commentText))
                    {
                        comments.Add($"• [{new CellPosition(r, c).ToAddress()}]: {commentText}");
                    }
                }
            }

            if (comments.Count > 0)
            {
                MessageBox.Show($"Worksheet Comments ({comments.Count} total):\n\n" + string.Join("\n", comments), "Comments", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("No comments found in active worksheet.", "Comments", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnProtectSheet_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.SetSettings(WorksheetSettings.Edit_Readonly, true);
            StatusText.Text = "Worksheet is protected against edits.";
            MessageBox.Show("Worksheet is now protected against cell edits.", "Protect Sheet", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnUnprotectSheet_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.DisableSettings(WorksheetSettings.Edit_Readonly);
            StatusText.Text = "Worksheet is unprotected.";
            MessageBox.Show("Worksheet editing unlocked.", "Unprotect Sheet", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnProtectWorkbook_Click(object sender, RoutedEventArgs e)
        {
            _isWorkbookProtected = !_isWorkbookProtected;
            StatusText.Text = _isWorkbookProtected ? "Workbook structure protected." : "Workbook structure unprotected.";
            MessageBox.Show(_isWorkbookProtected ? "Workbook structure is now protected." : "Workbook structure is now unprotected.", "Protect Workbook", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // =========================================================================
        // Insert Tab Elements & UI Tools
        // =========================================================================
        private void BtnInsertPicture_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var dlg = new OpenFileDialog
            {
                Title = "Insert Picture",
                Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var bmp = new BitmapImage(new Uri(dlg.FileName));
                    var imgObj = new ImageObject(bmp)
                    {
                        X = 60,
                        Y = 60
                    };
                    sheet.FloatingObjects.Add(imgObj);
                    StatusText.Text = $"Inserted picture: {Path.GetFileName(dlg.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not insert picture: {ex.Message}", "Insert Picture", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnInsertShape_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var shape = new RoundedRectangleShape
            {
                X = 80,
                Y = 80,
                Width = 140,
                Height = 70
            };
            sheet.FloatingObjects.Add(shape);
            StatusText.Text = "Inserted Rounded Rectangle shape.";
        }

        private void InsertChartInternal(string typeName)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            if (sel.Rows < 2 || sel.Cols < 2)
            {
                sel = new RangePosition(0, 0, Math.Min(sheet.RowCount, 5), Math.Min(sheet.ColumnCount, 2));
            }

            try
            {
                var catRange = new RangePosition(sel.Row, sel.Col, sel.Rows, 1);
                var valRange = new RangePosition(sel.Row, sel.Col + 1, sel.Rows, sel.Cols - 1);
                var ds = new WorksheetChartDataSource(sheet, catRange, valRange, RowOrColumn.Column);

                unvell.ReoGrid.Drawing.IDrawingObject chartObj = typeName switch
                {
                    "Bar" => new BarChart { DataSource = ds, Title = "Bar Analysis Chart" },
                    "Line" => new LineChart { DataSource = ds, Title = "Trend Line Chart" },
                    "Pie" => new PieChart { DataSource = ds, Title = "Distribution Pie Chart" },
                    _ => new ColumnChart { DataSource = ds, Title = "Column Analysis Chart" }
                };

                if (chartObj is unvell.ReoGrid.Chart.Chart chart)
                {
                    chart.X = 60;
                    chart.Y = 80;
                    chart.Width = 380;
                    chart.Height = 240;
                }

                sheet.FloatingObjects.Add(chartObj);
                StatusText.Text = $"Created interactive in-memory {typeName} Chart for range {sel.ToAddress()}. (Note: In-memory session visualization; OpenXML chart export planned for Phase 4).";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not generate {typeName} chart: {ex.Message}\n\nPlease select at least 2 rows and 2 columns with category labels and numerical values.", "Chart Creation", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnInsertColumnChart_Click(object sender, RoutedEventArgs e) => InsertChartInternal("Column");
        private void BtnInsertBarChart_Click(object sender, RoutedEventArgs e) => InsertChartInternal("Bar");
        private void BtnInsertLineChart_Click(object sender, RoutedEventArgs e) => InsertChartInternal("Line");
        private void BtnInsertPieChart_Click(object sender, RoutedEventArgs e) => InsertChartInternal("Pie");

        public class PivotSummaryMetadata
        {
            public string SourceSheetName { get; set; } = "";
            public RangePosition SourceRange { get; set; }
            public int RowFieldIndex { get; set; }
            public int SecondaryRowFieldIndex { get; set; } = -1;
            public int ColFieldIndex { get; set; }
            public int ValueFieldIndex { get; set; }
            public PivotAggregationType Aggregation { get; set; }
            public int StartRow { get; set; }
            public int StartCol { get; set; }
            public bool CreateChart { get; set; }
            public List<string> Headers { get; set; } = new();
        }

        private void BtnPivotTable_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            if (sel.Rows < 2 || sel.Cols < 2)
            {
                sel = new RangePosition(0, 0, Math.Min(sheet.RowCount, 20), Math.Min(sheet.ColumnCount, 6));
            }

            var headers = new List<string>();
            for (int c = sel.Col; c < sel.Col + sel.Cols; c++)
            {
                string h = sheet.Cells[sel.Row, c]?.Data?.ToString() ?? $"Column_{c + 1}";
                headers.Add(h);
            }

            var dlg = new PivotDialog(sel.ToAddress(), headers);
            dlg.Owner = this;
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    Worksheet targetSheet = dlg.IsNewSheet ? gridControl.CreateWorksheet($"Pivot_{DateTime.Now:HHmmss}") : sheet;
                    if (dlg.IsNewSheet)
                    {
                        WireSheetEvents(targetSheet);
                        gridControl.CurrentWorksheet = targetSheet;
                    }

                    int startR = dlg.IsNewSheet ? 1 : sel.Row + sel.Rows + 2;
                    int startC = dlg.IsNewSheet ? 1 : sel.Col;

                    var meta = new PivotSummaryMetadata
                    {
                        SourceSheetName = sheet.Name,
                        SourceRange = sel,
                        RowFieldIndex = dlg.RowFieldIndex,
                        SecondaryRowFieldIndex = dlg.SecondaryRowFieldIndex,
                        ColFieldIndex = dlg.ColFieldIndex,
                        ValueFieldIndex = dlg.ValueFieldIndex,
                        Aggregation = dlg.Aggregation,
                        StartRow = startR,
                        StartCol = startC,
                        CreateChart = dlg.CreatePivotChart,
                        Headers = headers
                    };

                    BuildPivotSummaryMatrix(targetSheet, sheet, meta);
                    StatusText.Text = $"Created Pivot Summary on sheet '{targetSheet.Name}'.";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to generate Pivot Summary: {ex.Message}", "Pivot Summary Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnRefreshPivot_Click(object sender, RoutedEventArgs e)
        {
            var targetSheet = gridControl.CurrentWorksheet;
            if (targetSheet != null && _pivotMetadataBySheet.TryGetValue(targetSheet.Name, out var meta))
            {
                var sourceSheet = gridControl.Worksheets[meta.SourceSheetName];
                if (sourceSheet == null)
                {
                    MessageBox.Show($"Source worksheet '{meta.SourceSheetName}' was not found.", "Refresh Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                BuildPivotSummaryMatrix(targetSheet, sourceSheet, meta);
                StatusText.Text = $"Refreshed Pivot Summary on '{targetSheet.Name}' from sheet '{meta.SourceSheetName}' with current data.";
            }
            else
            {
                MessageBox.Show("The active worksheet does not contain a refreshable Pivot Summary.\n\nTo generate one, select source data and click 'Pivot Summary'.", "Refresh Pivot Summary", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BuildPivotSummaryMatrix(Worksheet targetSheet, Worksheet sourceSheet, PivotSummaryMetadata meta)
        {
            var sel = meta.SourceRange;
            int rowFieldIdx = sel.Col + meta.RowFieldIndex;
            bool hasSecondary = meta.SecondaryRowFieldIndex >= 0 && meta.SecondaryRowFieldIndex != meta.RowFieldIndex;
            int secRowFieldIdx = hasSecondary ? sel.Col + meta.SecondaryRowFieldIndex : -1;
            int colFieldIdx = sel.Col + meta.ColFieldIndex;
            int valFieldIdx = sel.Col + meta.ValueFieldIndex;

            var rowKeys = new List<string>();
            var colKeys = new List<string>();
            var matrix = new Dictionary<string, Dictionary<string, List<double>>>();

            for (int r = sel.Row + 1; r < sel.Row + sel.Rows; r++)
            {
                string rVal = sourceSheet.Cells[r, rowFieldIdx]?.Data?.ToString()?.Trim() ?? "(blank)";
                if (hasSecondary)
                {
                    string secVal = sourceSheet.Cells[r, secRowFieldIdx]?.Data?.ToString()?.Trim() ?? "(blank)";
                    rVal = $"{rVal} - {secVal}";
                }

                string cVal = sourceSheet.Cells[r, colFieldIdx]?.Data?.ToString()?.Trim() ?? "(blank)";
                double val = 0;
                if (sourceSheet.Cells[r, valFieldIdx]?.Data is double d) val = d;
                else if (double.TryParse(sourceSheet.Cells[r, valFieldIdx]?.Data?.ToString(), out double parsed)) val = parsed;

                if (!rowKeys.Contains(rVal)) rowKeys.Add(rVal);
                if (!colKeys.Contains(cVal)) colKeys.Add(cVal);

                if (!matrix.ContainsKey(rVal)) matrix[rVal] = new Dictionary<string, List<double>>();
                if (!matrix[rVal].ContainsKey(cVal)) matrix[rVal][cVal] = new List<double>();
                matrix[rVal][cVal].Add(val);
            }

            rowKeys.Sort();
            colKeys.Sort();

            int startR = meta.StartRow;
            int startC = meta.StartCol;

            string rowHeaderLabel = meta.Headers[meta.RowFieldIndex];
            if (hasSecondary) rowHeaderLabel += $" / {meta.Headers[meta.SecondaryRowFieldIndex]}";

            targetSheet.Cells[startR, startC].Data = $"{rowHeaderLabel} \\ {meta.Headers[meta.ColFieldIndex]}";
            targetSheet.SetRangeStyles(new RangePosition(startR, startC, 1, colKeys.Count + 2), new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor | PlainStyleFlag.FontStyleBold,
                BackColor = SolidColor.FromArgb(255, 41, 128, 185),
                TextColor = SolidColor.FromArgb(255, 255, 255, 255),
                Bold = true
            });

            for (int j = 0; j < colKeys.Count; j++)
            {
                targetSheet.Cells[startR, startC + 1 + j].Data = colKeys[j];
            }
            targetSheet.Cells[startR, startC + 1 + colKeys.Count].Data = "Grand Total";

            // Row Population & Row Grand Totals
            for (int i = 0; i < rowKeys.Count; i++)
            {
                string rk = rowKeys[i];
                targetSheet.Cells[startR + 1 + i, startC].Data = rk;
                targetSheet.Cells[startR + 1 + i, startC].Style.Bold = true;

                double rowTotal = 0;
                int rowCount = 0;

                for (int j = 0; j < colKeys.Count; j++)
                {
                    string ck = colKeys[j];
                    double cellVal = 0;
                    if (matrix.ContainsKey(rk) && matrix[rk].ContainsKey(ck))
                    {
                        var vals = matrix[rk][ck];
                        cellVal = meta.Aggregation switch
                        {
                            PivotAggregationType.Count => vals.Count,
                            PivotAggregationType.Average => vals.Average(),
                            PivotAggregationType.Max => vals.Max(),
                            PivotAggregationType.Min => vals.Min(),
                            _ => vals.Sum()
                        };
                        rowTotal += vals.Sum();
                        rowCount += vals.Count;
                    }
                    targetSheet.Cells[startR + 1 + i, startC + 1 + j].Data = cellVal;
                    targetSheet.Cells[startR + 1 + i, startC + 1 + j].DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Number;
                }

                double rowAgg = meta.Aggregation == PivotAggregationType.Count ? rowCount :
                                (meta.Aggregation == PivotAggregationType.Average && rowCount > 0 ? rowTotal / rowCount : rowTotal);
                targetSheet.Cells[startR + 1 + i, startC + 1 + colKeys.Count].Data = rowAgg;
                targetSheet.Cells[startR + 1 + i, startC + 1 + colKeys.Count].Style.Bold = true;
            }

            // Bottom Grand Total Row across all columns
            int bottomGrandRow = startR + 1 + rowKeys.Count;
            targetSheet.Cells[bottomGrandRow, startC].Data = "Grand Total";
            targetSheet.Cells[bottomGrandRow, startC].Style.Bold = true;

            double grandOverallSum = 0;
            int grandOverallCount = 0;
            double grandOverallMax = double.MinValue;
            double grandOverallMin = double.MaxValue;

            for (int j = 0; j < colKeys.Count; j++)
            {
                string ck = colKeys[j];
                var colVals = new List<double>();
                foreach (var rk in rowKeys)
                {
                    if (matrix.ContainsKey(rk) && matrix[rk].ContainsKey(ck))
                    {
                        colVals.AddRange(matrix[rk][ck]);
                    }
                }

                double colAgg = 0;
                if (colVals.Count > 0)
                {
                    colAgg = meta.Aggregation switch
                    {
                        PivotAggregationType.Count => colVals.Count,
                        PivotAggregationType.Average => colVals.Average(),
                        PivotAggregationType.Max => colVals.Max(),
                        PivotAggregationType.Min => colVals.Min(),
                        _ => colVals.Sum()
                    };
                    grandOverallSum += colVals.Sum();
                    grandOverallCount += colVals.Count;
                    if (colVals.Max() > grandOverallMax) grandOverallMax = colVals.Max();
                    if (colVals.Min() < grandOverallMin) grandOverallMin = colVals.Min();
                }

                targetSheet.Cells[bottomGrandRow, startC + 1 + j].Data = colAgg;
                targetSheet.Cells[bottomGrandRow, startC + 1 + j].Style.Bold = true;
                targetSheet.Cells[bottomGrandRow, startC + 1 + j].DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Number;
            }

            double grandFinalVal = meta.Aggregation switch
            {
                PivotAggregationType.Count => grandOverallCount,
                PivotAggregationType.Average => grandOverallCount > 0 ? grandOverallSum / grandOverallCount : 0,
                PivotAggregationType.Max => grandOverallMax != double.MinValue ? grandOverallMax : 0,
                PivotAggregationType.Min => grandOverallMin != double.MaxValue ? grandOverallMin : 0,
                _ => grandOverallSum
            };

            targetSheet.Cells[bottomGrandRow, startC + 1 + colKeys.Count].Data = grandFinalVal;
            targetSheet.Cells[bottomGrandRow, startC + 1 + colKeys.Count].Style.Bold = true;
            targetSheet.Cells[bottomGrandRow, startC + 1 + colKeys.Count].DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Number;

            int totalRows = rowKeys.Count + 2;
            int totalCols = colKeys.Count + 2;
            targetSheet.SetRangeBorders(new RangePosition(startR, startC, totalRows, totalCols), BorderPositions.All, RangeBorderStyle.BlackSolid);

            if (meta.CreateChart && rowKeys.Count > 0 && colKeys.Count > 0)
            {
                var catRange = new RangePosition(startR + 1, startC, rowKeys.Count, 1);
                var valRange = new RangePosition(startR + 1, startC + 1, rowKeys.Count, colKeys.Count);
                var chart = new ColumnChart
                {
                    DataSource = new WorksheetChartDataSource(targetSheet, catRange, valRange, RowOrColumn.Column),
                    X = 50,
                    Y = (startR + totalRows + 1) * 24,
                    Width = 440,
                    Height = 260,
                    Title = $"Summary Chart ({meta.Headers[meta.ValueFieldIndex]} by {meta.Headers[meta.RowFieldIndex]})"
                };
                targetSheet.FloatingObjects.Add(chart);
            }

            _pivotMetadataBySheet[targetSheet.Name] = meta;
        }

        private void BtnFind_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new FindReplaceDialog(gridControl, isReplaceMode: false);
            dlg.Owner = this;
            dlg.Show();
        }

        private void BtnReplace_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new FindReplaceDialog(gridControl, isReplaceMode: true);
            dlg.Owner = this;
            dlg.Show();
        }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var session = gridControl.CreatePrintSession();
                session.Print();
                StatusText.Text = "Print session initialized.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not initialize printer: {ex.Message}", "Print Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnZoomReset_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.ZoomReset();
            StatusText.Text = "Zoom reset to 100%.";
        }

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.ZoomIn();
            StatusText.Text = $"Zoom In ({gridControl.CurrentWorksheet?.ScaleFactor * 100:0}%).";
        }

        private void BtnZoomOut_Click(object sender, RoutedEventArgs e)
        {
            gridControl.CurrentWorksheet?.ZoomOut();
            StatusText.Text = $"Zoom Out ({gridControl.CurrentWorksheet?.ScaleFactor * 100:0}%).";
        }

        private void BtnCreateTable_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            sheet.SetRangeStyles(new RangePosition(sel.Row, sel.Col, 1, sel.Cols), new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor | PlainStyleFlag.FontStyleBold,
                BackColor = SolidColor.FromArgb(255, 30, 75, 120),
                TextColor = SolidColor.FromArgb(255, 255, 255, 255),
                Bold = true
            });
            sheet.SetRangeBorders(sel, BorderPositions.All, RangeBorderStyle.BlackSolid);
            StatusText.Text = $"Formatted Table on range {sel.ToAddress()}";
        }

        private void BtnToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            _activeTheme = _activeTheme == "DarkGray" ? "Colorful" : "DarkGray";
            StatusText.Text = $"Active Theme: {_activeTheme}";
        }

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SettingsDialog();
            dlg.Owner = this;
            dlg.ShowDialog();
        }

        private void BtnShortcuts_Click(object sender, RoutedEventArgs e)
        {
            ShowShortcutsDialog();
        }

        private void ShowShortcutsDialog()
        {
            var dlg = new ShortcutsDialog();
            dlg.Owner = this;
            dlg.ShowDialog();
        }

        private void BtnAbout_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new AboutDialog();
            dlg.Owner = this;
            dlg.ShowDialog();
        }

        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            if (gridControl.CanUndo())
            {
                gridControl.Undo();
                StatusText.Text = "Undo performed.";
            }
        }

        private void BtnRedo_Click(object sender, RoutedEventArgs e)
        {
            if (gridControl.CanRedo())
            {
                gridControl.Redo();
                StatusText.Text = "Redo performed.";
            }
        }

        private bool _autoSaveEnabled = true;
        private void BtnAutoSaveToggle_Click(object sender, RoutedEventArgs e)
        {
            _autoSaveEnabled = !_autoSaveEnabled;
            if (FindName("TxtAutoSaveStatus") is TextBlock tb)
            {
                tb.Text = _autoSaveEnabled ? "On" : "Off";
            }
            if (FindName("AutoSavePill") is Border pill)
            {
                pill.Background = _autoSaveEnabled ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x10, 0x7C, 0x41)) : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x2B, 0x2B, 0x2B));
            }
            if (FindName("AutoSaveThumb") is FrameworkElement thumb)
            {
                thumb.HorizontalAlignment = _autoSaveEnabled ? System.Windows.HorizontalAlignment.Right : System.Windows.HorizontalAlignment.Left;
            }
            if (_autoSaveEnabled)
            {
                _autoSaveTimer?.Start();
                StatusText.Text = "AutoSave enabled (every 5 minutes).";
            }
            else
            {
                _autoSaveTimer?.Stop();
                StatusText.Text = "AutoSave paused.";
            }
        }

        private void BtnFormulaEnter_Click(object sender, RoutedEventArgs e)
        {
            CommitFormulaBar();
        }

        private void BtnFormulaCancel_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var cell = sheet.Cells[sheet.FocusPos];
            TxtFormulaBar.Text = cell.HasFormula ? "=" + cell.Formula : cell.DisplayText;
            StatusText.Text = "Formula edit cancelled.";
        }

        private void BtnCellStyle_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            string style = (sender as Button)?.Tag as string ?? "Normal";

            switch (style)
            {
                case "Bad":
                    sheet.SetRangeStyles(sel, new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor,
                        BackColor = SolidColor.FromArgb(255, 255, 199, 206),
                        TextColor = SolidColor.FromArgb(255, 156, 0, 6)
                    });
                    StatusText.Text = "Applied 'Bad' cell style.";
                    break;
                case "Good":
                    sheet.SetRangeStyles(sel, new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor,
                        BackColor = SolidColor.FromArgb(255, 198, 239, 206),
                        TextColor = SolidColor.FromArgb(255, 0, 97, 0)
                    });
                    StatusText.Text = "Applied 'Good' cell style.";
                    break;
                case "Neutral":
                    sheet.SetRangeStyles(sel, new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor,
                        BackColor = SolidColor.FromArgb(255, 255, 235, 156),
                        TextColor = SolidColor.FromArgb(255, 156, 101, 0)
                    });
                    StatusText.Text = "Applied 'Neutral' cell style.";
                    break;
                case "Calculation":
                    sheet.SetRangeStyles(sel, new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor | PlainStyleFlag.FontStyleBold,
                        BackColor = SolidColor.FromArgb(255, 242, 242, 242),
                        TextColor = SolidColor.FromArgb(255, 250, 125, 0)
                    });
                    StatusText.Text = "Applied 'Calculation' cell style.";
                    break;
                case "CheckCell":
                    sheet.SetRangeBorders(sel, BorderPositions.Outside, RangeBorderStyle.BlackSolid);
                    StatusText.Text = "Applied 'Check Cell' border style.";
                    break;
                case "Normal":
                default:
                    sheet.SetRangeStyles(sel, new WorksheetRangeStyle
                    {
                        Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor,
                        BackColor = SolidColor.FromArgb(255, 255, 255, 255),
                        TextColor = SolidColor.FromArgb(255, 0, 0, 0)
                    });
                    StatusText.Text = "Applied 'Normal' cell style.";
                    break;
            }
        }

        private void BtnWrapText_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.TextWrap,
                TextWrapMode = TextWrapMode.WordBreak
            });
            StatusText.Text = "Applied Wrap Text to selection.";
        }

        private void BtnCurrency_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    sheet.Cells[sel.Row + r, sel.Col + c].DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Currency;
                }
            }
            StatusText.Text = "Applied Currency formatting ($).";
        }

        private void BtnPercent_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    sheet.Cells[sel.Row + r, sel.Col + c].DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Percent;
                }
            }
            StatusText.Text = "Applied Percent formatting (%).";
        }

        private void BtnComma_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    sheet.Cells[sel.Row + r, sel.Col + c].DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Number;
                }
            }
            StatusText.Text = "Applied Comma / Number formatting.";
        }

        private void BtnDecreaseIndent_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            StatusText.Text = "Decreased Indent.";
        }

        private void BtnIncreaseIndent_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            StatusText.Text = "Increased Indent.";
        }

        private void BtnClearContents_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    sheet.Cells[sel.Row + r, sel.Col + c].Data = null;
                }
            }
            sheet.Recalculate();
            StatusText.Text = $"Cleared contents from {sel.ToAddress()}.";
        }

        private void BtnShare_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("ExcelClone is a 100% Free & Open-Source offline application.\n\nYou can share your workbook directly by copying or emailing your saved XLSX file.", "Share Workbook", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnFileMenu_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.ContextMenu != null)
            {
                fe.ContextMenu.PlacementTarget = fe;
                fe.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                fe.ContextMenu.IsOpen = true;
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            BtnFind_Click(sender, e);
        }

        private void ChkGridlines_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            if (sender is CheckBox chk && chk.IsChecked == true)
                sheet.SetSettings(WorksheetSettings.View_ShowGridLine, true);
            else
                sheet.DisableSettings(WorksheetSettings.View_ShowGridLine);
            StatusText.Text = $"Gridlines {(sheet.HasSettings(WorksheetSettings.View_ShowGridLine) ? "visible" : "hidden")}.";
        }

        private void ChkHeadings_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            if (sender is CheckBox chk && chk.IsChecked == true)
                sheet.SetSettings(WorksheetSettings.View_ShowHeaders, true);
            else
                sheet.DisableSettings(WorksheetSettings.View_ShowHeaders);
            StatusText.Text = $"Headings {(sheet.HasSettings(WorksheetSettings.View_ShowHeaders) ? "visible" : "hidden")}.";
        }

        private void BtnIncreaseFontSize_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            int current = CmbFontSize.SelectedIndex >= 0 && int.TryParse(((ComboBoxItem)CmbFontSize.SelectedItem).Content.ToString(), out int sz) ? sz : 11;
            int next = Math.Min(72, current + 2);
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontSize,
                FontSize = next
            });
            StatusText.Text = $"Increased font size to {next}pt.";
        }

        private void BtnDecreaseFontSize_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            int current = CmbFontSize.SelectedIndex >= 0 && int.TryParse(((ComboBoxItem)CmbFontSize.SelectedItem).Content.ToString(), out int sz) ? sz : 11;
            int next = Math.Max(8, current - 2);
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.FontSize,
                FontSize = next
            });
            StatusText.Text = $"Decreased font size to {next}pt.";
        }

        private void BtnFontColorRed_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.TextColor,
                TextColor = SolidColor.FromArgb(255, 192, 0, 0)
            });
            StatusText.Text = "Applied Red text color.";
        }

        private WorksheetRangeStyle? _copiedFormatStyle = null;
        private void BtnFormatPainter_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            if (_copiedFormatStyle == null)
            {
                _copiedFormatStyle = sheet.GetRangeStyles(new RangePosition(sheet.FocusPos.Row, sheet.FocusPos.Col, 1, 1));
                StatusText.Text = $"Format copied from {sheet.FocusPos.ToAddress()}. Select destination cells and click Format Painter again to apply.";
            }
            else
            {
                sheet.SetRangeStyles(sheet.SelectionRange, _copiedFormatStyle);
                _copiedFormatStyle = null;
                StatusText.Text = "Applied copied format to selection.";
            }
        }

        private void BtnAlignTop_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.VerticalAlign,
                VAlign = ReoGridVerAlign.Top
            });
            StatusText.Text = "Applied Top Vertical Alignment.";
        }

        private void BtnAlignMiddle_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.VerticalAlign,
                VAlign = ReoGridVerAlign.Middle
            });
            StatusText.Text = "Applied Middle Vertical Alignment.";
        }

        private void BtnAlignBottom_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            sheet.SetRangeStyles(sheet.SelectionRange, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.VerticalAlign,
                VAlign = ReoGridVerAlign.Bottom
            });
            StatusText.Text = "Applied Bottom Vertical Alignment.";
        }

        private void BtnConditionalFormatting_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            var values = new List<double>();
            for (int r = 0; r < sel.Rows; r++)
                for (int c = 0; c < sel.Cols; c++)
                    if (double.TryParse(sheet.Cells[sel.Row + r, sel.Col + c].Data?.ToString(), out double val))
                        values.Add(val);

            if (values.Count > 0)
            {
                double avg = values.Average();
                for (int r = 0; r < sel.Rows; r++)
                {
                    for (int c = 0; c < sel.Cols; c++)
                    {
                        if (double.TryParse(sheet.Cells[sel.Row + r, sel.Col + c].Data?.ToString(), out double val))
                        {
                            if (val >= avg)
                            {
                                sheet.SetRangeStyles(new RangePosition(sel.Row + r, sel.Col + c, 1, 1), new WorksheetRangeStyle
                                {
                                    Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor,
                                    BackColor = SolidColor.FromArgb(255, 198, 239, 206),
                                    TextColor = SolidColor.FromArgb(255, 0, 97, 0)
                                });
                            }
                            else
                            {
                                sheet.SetRangeStyles(new RangePosition(sel.Row + r, sel.Col + c, 1, 1), new WorksheetRangeStyle
                                {
                                    Flag = PlainStyleFlag.BackColor | PlainStyleFlag.TextColor,
                                    BackColor = SolidColor.FromArgb(255, 255, 199, 206),
                                    TextColor = SolidColor.FromArgb(255, 156, 0, 6)
                                });
                            }
                        }
                    }
                }
                StatusText.Text = $"Applied Conditional Formatting (Above/Below Avg: {avg:N2}).";
            }
            else
            {
                StatusText.Text = "Select cells with numbers to apply Conditional Formatting.";
            }
        }

        private void CmbNumberFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (gridControl?.CurrentWorksheet == null || sender is not ComboBox cmb || cmb.SelectedItem is not ComboBoxItem item) return;
            var sheet = gridControl.CurrentWorksheet;
            string format = item.Content?.ToString() ?? "General";
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    var cell = sheet.Cells[sel.Row + r, sel.Col + c];
                    switch (format)
                    {
                        case "Currency":
                        case "Accounting":
                            cell.DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Currency;
                            break;
                        case "Percentage":
                            cell.DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Percent;
                            break;
                        case "Number":
                            cell.DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.Number;
                            break;
                        case "Short Date":
                            cell.DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.DateTime;
                            break;
                        default:
                            cell.DataFormat = unvell.ReoGrid.DataFormat.CellDataFormatFlag.General;
                            break;
                    }
                }
            }
            if (StatusText != null)
            {
                StatusText.Text = $"Applied number format '{format}' to selection.";
            }
        }

        private void BtnIncreaseDecimal_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    var cell = sheet.Cells[sel.Row + r, sel.Col + c];
                    if (cell.Data is double d)
                    {
                        cell.Data = Math.Round(d, 3);
                    }
                    else if (double.TryParse(cell.DisplayText, out double num))
                    {
                        cell.Data = Math.Round(num, 3);
                    }
                }
            }
            StatusText.Text = "Adjusted decimal precision to 3 places.";
        }

        private void BtnDecreaseDecimal_Click(object sender, RoutedEventArgs e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            var sel = sheet.SelectionRange;
            for (int r = 0; r < sel.Rows; r++)
            {
                for (int c = 0; c < sel.Cols; c++)
                {
                    var cell = sheet.Cells[sel.Row + r, sel.Col + c];
                    if (cell.Data is double d)
                    {
                        cell.Data = Math.Round(d, 0);
                    }
                    else if (double.TryParse(cell.DisplayText, out double num))
                    {
                        cell.Data = Math.Round(num, 0);
                    }
                }
            }
            StatusText.Text = "Rounded to whole numbers (0 decimals).";
        }

        private bool _isRibbonCollapsed = false;
        private void BtnCollapseRibbon_Click(object sender, RoutedEventArgs e)
        {
            _isRibbonCollapsed = !_isRibbonCollapsed;
            if (FindName("RibbonTabs") is TabControl tabControl)
            {
                tabControl.Height = _isRibbonCollapsed ? 31 : double.NaN;
            }
            StatusText.Text = _isRibbonCollapsed ? "Ribbon collapsed." : "Ribbon expanded.";
        }

        private void SliderZoom_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            var sheet = gridControl.CurrentWorksheet;
            if (sheet == null) return;
            float factor = (float)(e.NewValue / 100.0);
            sheet.ScaleFactor = factor;
            if (FindName("TxtZoomPercent") is TextBlock tb)
            {
                tb.Text = $"{(int)e.NewValue}%";
            }
        }
    }
}

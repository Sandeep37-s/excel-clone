using System;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;
using Xunit;
using ExcelClone;

namespace ExcelClone.Tests
{
    public class AllExcelButtonsAndFeaturesTests
    {
        [Fact]
        public void VerifyAllHomeRibbonButtonsAndFormatting()
        {
            var t = new Thread(() =>
            {
                var win = new MainWindow();
                var grid = (Grid)win.Content;
                var gridControl = (ReoGridControl)win.FindName("gridControl");
                Assert.NotNull(gridControl);
                var sheet = gridControl.CurrentWorksheet;
                Assert.NotNull(sheet);

                // 1. Enter test numbers in cells
                sheet.Cells["A1"].Data = 10.0;
                sheet.Cells["A2"].Data = 20.0;
                sheet.Cells["A3"].Data = 30.0;

                // 2. Select A1:A3 and test Font size increase
                sheet.SelectRange(new RangePosition(0, 0, 3, 1));
                var cmbFontSize = (ComboBox)win.FindName("CmbFontSize");
                cmbFontSize.SelectedIndex = 2; // 11pt
                
                var btnIncrease = new Button();
                // Call font size increase
                var miIncrease = typeof(MainWindow).GetMethod("BtnIncreaseFontSize_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miIncrease);
                miIncrease.Invoke(win, new object[] { btnIncrease, new RoutedEventArgs() });

                // Verify font size changed on selection
                var styleA1 = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(13, styleA1.FontSize);

                // 3. Test Red Font Color
                var miRed = typeof(MainWindow).GetMethod("BtnFontColorRed_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miRed);
                miRed.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                styleA1 = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(192, styleA1.TextColor.R);

                // 4. Test Vertical Alignments
                var miAlignMiddle = typeof(MainWindow).GetMethod("BtnAlignMiddle_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miAlignMiddle);
                miAlignMiddle.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                styleA1 = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(ReoGridVerAlign.Middle, styleA1.VAlign);

                // 5. Test Horizontal Align Center
                var miAlignCenter = typeof(MainWindow).GetMethod("BtnAlignCenter_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miAlignCenter);
                miAlignCenter.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                styleA1 = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(ReoGridHorAlign.Center, styleA1.HAlign);

                // 6. Test Currency formatting
                var miCurrency = typeof(MainWindow).GetMethod("BtnCurrency_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miCurrency);
                miCurrency.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                Assert.Equal(unvell.ReoGrid.DataFormat.CellDataFormatFlag.Currency, sheet.Cells["A1"].DataFormat);

                // 7. Test Percentage formatting
                var miPercent = typeof(MainWindow).GetMethod("BtnPercent_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miPercent);
                miPercent.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                Assert.Equal(unvell.ReoGrid.DataFormat.CellDataFormatFlag.Percent, sheet.Cells["A1"].DataFormat);

                // 8. Test Conditional Formatting (Average of 10, 20, 30 is 20: A3 is above, A1 is below)
                var miCond = typeof(MainWindow).GetMethod("BtnConditionalFormatting_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miCond);
                miCond.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                var styleA3 = sheet.GetRangeStyles(new RangePosition(2, 0, 1, 1));
                Assert.Equal(198, styleA3.BackColor.R); // Light green #C6EFCE

                var styleA1_cond = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(255, styleA1_cond.BackColor.R); // Light pink #FFC7CE

                // 9. Test Cell Styles Presets
                var miStyle = typeof(MainWindow).GetMethod("BtnCellStyle_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miStyle);

                // Apply Good
                var btnGood = new Button { Tag = "Good" };
                miStyle.Invoke(win, new object[] { btnGood, new RoutedEventArgs() });
                var styleGood = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(198, styleGood.BackColor.R);

                // Apply Bad
                var btnBad = new Button { Tag = "Bad" };
                miStyle.Invoke(win, new object[] { btnBad, new RoutedEventArgs() });
                var styleBad = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(255, styleBad.BackColor.R);
                Assert.Equal(156, styleBad.TextColor.R);

                // Apply Normal
                var btnNormal = new Button { Tag = "Normal" };
                miStyle.Invoke(win, new object[] { btnNormal, new RoutedEventArgs() });
                var styleNormal = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(255, styleNormal.BackColor.R);
                Assert.Equal(0, styleNormal.TextColor.R);

                // 10. Test AutoSave Toggle
                var txtAutoSave = (TextBlock)win.FindName("TxtAutoSaveStatus");
                var miAutoSave = typeof(MainWindow).GetMethod("BtnAutoSaveToggle_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miAutoSave);
                miAutoSave.Invoke(win, new object[] { new Button(), new RoutedEventArgs() }); // toggles to Off
                Assert.Equal("Off", txtAutoSave.Text);
                miAutoSave.Invoke(win, new object[] { new Button(), new RoutedEventArgs() }); // toggles back to On
                Assert.Equal("On", txtAutoSave.Text);

                // 11. Test Formula Bar Enter and Recalculate
                sheet.FocusPos = new CellPosition("B1");
                var txtFormulaBar = (TextBox)win.FindName("TxtFormulaBar");
                txtFormulaBar.Text = "=SUM(A1:A3)";
                var miFormulaEnter = typeof(MainWindow).GetMethod("BtnFormulaEnter_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miFormulaEnter);
                miFormulaEnter.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                Assert.Equal("SUM(A1:A3)", sheet.Cells["B1"].Formula);
                Assert.Equal(60.0, Convert.ToDouble(sheet.Cells["B1"].Data));

                // 12. Test Clear Contents
                sheet.SelectRange(new RangePosition(0, 0, 1, 1)); // Select A1
                var miClear = typeof(MainWindow).GetMethod("BtnClearContents_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miClear);
                miClear.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.Null(sheet.Cells["A1"].Data);

                // 13. Test Gridlines and Headings settings
                var miGridlines = typeof(MainWindow).GetMethod("ChkGridlines_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miGridlines);
                var chkGridlines = new CheckBox { IsChecked = false };
                miGridlines.Invoke(win, new object[] { chkGridlines, new RoutedEventArgs() });
                Assert.False(sheet.HasSettings(WorksheetSettings.View_ShowGridLine));

                chkGridlines.IsChecked = true;
                miGridlines.Invoke(win, new object[] { chkGridlines, new RoutedEventArgs() });
                Assert.True(sheet.HasSettings(WorksheetSettings.View_ShowGridLine));
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            Assert.True(t.Join(10000));
        }

        [Fact]
        public void VerifyAllOtherRibbonTabsAndActions()
        {
            var t = new Thread(() =>
            {
                var win = new MainWindow();
                var gridControl = (ReoGridControl)win.FindName("gridControl");
                Assert.NotNull(gridControl);
                var sheet = gridControl.CurrentWorksheet;
                Assert.NotNull(sheet);

                // 1. Data Setup
                sheet.Cells["A1"].Data = "Banana";
                sheet.Cells["A2"].Data = "Apple";
                sheet.Cells["A3"].Data = "Cherry";
                sheet.Cells["B1"].Data = 30.0;
                sheet.Cells["B2"].Data = 10.0;
                sheet.Cells["B3"].Data = 20.0;

                // 2. Test Sort A to Z
                sheet.SelectRange(new RangePosition(0, 0, 3, 2));
                var miSortAsc = typeof(MainWindow).GetMethod("BtnSortAsc_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miSortAsc);
                miSortAsc.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                Assert.Equal("Apple", sheet.Cells["A1"].Data?.ToString());
                Assert.Equal("Banana", sheet.Cells["A2"].Data?.ToString());
                Assert.Equal("Cherry", sheet.Cells["A3"].Data?.ToString());

                // 3. Test Sort Z to A
                var miSortDesc = typeof(MainWindow).GetMethod("BtnSortDesc_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miSortDesc);
                miSortDesc.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                Assert.Equal("Cherry", sheet.Cells["A1"].Data?.ToString());
                Assert.Equal("Banana", sheet.Cells["A2"].Data?.ToString());
                Assert.Equal("Apple", sheet.Cells["A3"].Data?.ToString());

                // 4. Test Table Creation
                sheet.SelectRange(new RangePosition(0, 0, 3, 2));
                var miTable = typeof(MainWindow).GetMethod("BtnCreateTable_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miTable);
                miTable.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });

                var headerStyle = sheet.GetRangeStyles(new RangePosition(0, 0, 1, 1));
                Assert.Equal(30, headerStyle.BackColor.R); // SolidColor.FromArgb(255, 30, 75, 120)
                Assert.Equal(75, headerStyle.BackColor.G);

                // 5. Test Add Sheet
                int initialSheets = gridControl.Worksheets.Count;
                var miAddSheet = typeof(MainWindow).GetMethod("BtnAddSheet_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miAddSheet);
                miAddSheet.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.Equal(initialSheets + 1, gridControl.Worksheets.Count);

                // 6. Test Freeze Panes & Unfreeze Panes on active worksheet
                var activeSheet = gridControl.CurrentWorksheet;
                Assert.NotNull(activeSheet);
                activeSheet.FreezeToCell(1, 0);
                Assert.True(activeSheet.IsFrozen);

                var miUnfreeze = typeof(MainWindow).GetMethod("BtnUnfreezePanes_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miUnfreeze);
                miUnfreeze.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.False(activeSheet.IsFrozen);

                var miFreeze = typeof(MainWindow).GetMethod("BtnFreezePanes_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miFreeze);
                miFreeze.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.True(activeSheet.IsFrozen);

                // 7. Test Chart Insertions (Column, Bar, Line, Pie)
                activeSheet.Cells["A1"].Data = "North";
                activeSheet.Cells["A2"].Data = "South";
                activeSheet.Cells["B1"].Data = 100.0;
                activeSheet.Cells["B2"].Data = 200.0;
                activeSheet.SelectRange(new RangePosition(0, 0, 2, 2));
                var miColChart = typeof(MainWindow).GetMethod("BtnInsertColumnChart_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miColChart);
                miColChart.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.NotEmpty(activeSheet.FloatingObjects);

                // 8. Test Draw Inking Mode
                var drawingCanvas = (InkCanvas)win.FindName("drawingCanvas");
                Assert.NotNull(drawingCanvas);
                var miBlackPen = typeof(MainWindow).GetMethod("BtnBlackPen_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miBlackPen);
                miBlackPen.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.Equal(Visibility.Visible, drawingCanvas.Visibility);

                var miDrawSelect = typeof(MainWindow).GetMethod("BtnDrawSelect_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miDrawSelect);
                miDrawSelect.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.Equal(Visibility.Collapsed, drawingCanvas.Visibility);

                // 9. Test Zoom In, Zoom Out, Zoom Reset
                var miZoomIn = typeof(MainWindow).GetMethod("BtnZoomIn_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miZoomIn);
                miZoomIn.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.True(activeSheet.ScaleFactor > 1.0f);

                var miZoomReset = typeof(MainWindow).GetMethod("BtnZoomReset_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miZoomReset);
                miZoomReset.Invoke(win, new object[] { new Button(), new RoutedEventArgs() });
                Assert.Equal(1.0f, activeSheet.ScaleFactor);

                // 10. Test Theme Toggle
                var miTheme = typeof(MainWindow).GetMethod("BtnToggleTheme_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(miTheme);
                miTheme.Invoke(win, new object[] { new Button(), new RoutedEventArgs() }); // toggles to Colorful
                miTheme.Invoke(win, new object[] { new Button(), new RoutedEventArgs() }); // toggles back to DarkGray
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            Assert.True(t.Join(10000));
        }
    }
}

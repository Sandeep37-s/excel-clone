using System;
using System.Collections.Generic;
using System.Windows;
using unvell.ReoGrid;

namespace ExcelClone
{
    public partial class FindReplaceDialog : Window
    {
        private readonly ReoGridControl _gridControl;
        private int _lastFoundRow = -1;
        private int _lastFoundCol = -1;
        private Worksheet? _lastFoundSheet = null;

        public FindReplaceDialog(ReoGridControl gridControl, bool isReplaceMode = false)
        {
            InitializeComponent();
            _gridControl = gridControl;
            if (!isReplaceMode)
            {
                TxtReplace.IsEnabled = false;
                BtnReplace.IsEnabled = false;
                BtnReplaceAll.IsEnabled = false;
            }
            Loaded += (s, e) => TxtFind.Focus();
        }

        private void BtnFindNext_Click(object sender, RoutedEventArgs e)
        {
            string query = TxtFind.Text;
            if (string.IsNullOrEmpty(query))
            {
                TxtStatus.Text = "Please enter text to search for.";
                return;
            }

            bool matchCase = ChkMatchCase.IsChecked == true;
            bool matchEntire = ChkMatchEntire.IsChecked == true;
            bool entireWb = RbWorkbook.IsChecked == true;

            IEnumerable<Worksheet> sheets = entireWb ? _gridControl.Worksheets : new[] { _gridControl.CurrentWorksheet };
            StringComparison comp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            var sheetList = new List<Worksheet>(sheets);
            int startSheetIdx = 0;
            if (_lastFoundSheet != null && entireWb)
            {
                startSheetIdx = sheetList.IndexOf(_lastFoundSheet);
                if (startSheetIdx < 0) startSheetIdx = 0;
            }

            for (int sIdx = 0; sIdx < sheetList.Count; sIdx++)
            {
                int currSheetIdx = (startSheetIdx + sIdx) % sheetList.Count;
                var ws = sheetList[currSheetIdx];
                if (ws == null) continue;

                int maxR = Math.Min(ws.RowCount, 500);
                int maxC = Math.Min(ws.ColumnCount, 100);

                int startR = (ws == _lastFoundSheet && _lastFoundRow >= 0) ? _lastFoundRow : 0;
                int startC = (ws == _lastFoundSheet && _lastFoundCol >= 0) ? _lastFoundCol + 1 : 0;

                for (int r = startR; r < maxR; r++)
                {
                    int cStart = (r == startR) ? startC : 0;
                    for (int c = cStart; c < maxC; c++)
                    {
                        var cell = ws.Cells[r, c];
                        string val = cell?.Data?.ToString() ?? "";
                        if (string.IsNullOrEmpty(val) && !string.IsNullOrEmpty(cell?.Formula))
                        {
                            val = "=" + cell.Formula;
                        }

                        bool match = false;
                        if (matchEntire)
                        {
                            match = string.Equals(val, query, comp);
                        }
                        else
                        {
                            match = val.IndexOf(query, comp) >= 0;
                        }

                        if (match)
                        {
                            _lastFoundRow = r;
                            _lastFoundCol = c;
                            _lastFoundSheet = ws;

                            if (_gridControl.CurrentWorksheet != ws)
                            {
                                _gridControl.CurrentWorksheet = ws;
                            }
                            ws.FocusPos = new CellPosition(r, c);
                            ws.SelectRange(new RangePosition(r, c, 1, 1));
                            TxtStatus.Text = $"Found '{query}' in {ws.Name}!{new CellPosition(r, c).ToAddress()}";
                            return;
                        }
                    }
                }
            }

            // If we reached here, reset search wrap
            if (_lastFoundRow >= 0)
            {
                _lastFoundRow = -1;
                _lastFoundCol = -1;
                _lastFoundSheet = null;
                TxtStatus.Text = $"Reached end of sheet/workbook. Searching again from top...";
                BtnFindNext_Click(sender, e);
                return;
            }

            TxtStatus.Text = $"Cannot find '{query}'.";
        }

        private void BtnReplace_Click(object sender, RoutedEventArgs e)
        {
            var ws = _gridControl.CurrentWorksheet;
            if (ws == null) return;
            var pos = ws.FocusPos;
            var cell = ws.Cells[pos];
            string val = cell?.Data?.ToString() ?? "";
            string query = TxtFind.Text;
            string replaceWith = TxtReplace.Text ?? "";

            bool matchCase = ChkMatchCase.IsChecked == true;
            StringComparison comp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            if (cell != null && !string.IsNullOrEmpty(query) && val.IndexOf(query, comp) >= 0)
            {
                string newVal = val.Replace(query, replaceWith, comp);
                if (double.TryParse(newVal, out double num))
                    cell.Data = num;
                else
                    cell.Data = newVal;

                TxtStatus.Text = $"Replaced 1 occurrence at {pos.ToAddress()}.";
            }

            BtnFindNext_Click(sender, e);
        }

        private void BtnReplaceAll_Click(object sender, RoutedEventArgs e)
        {
            string query = TxtFind.Text;
            string replaceWith = TxtReplace.Text ?? "";
            if (string.IsNullOrEmpty(query))
            {
                TxtStatus.Text = "Please specify text to find.";
                return;
            }

            bool matchCase = ChkMatchCase.IsChecked == true;
            bool matchEntire = ChkMatchEntire.IsChecked == true;
            bool entireWb = RbWorkbook.IsChecked == true;
            StringComparison comp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            IEnumerable<Worksheet> sheets = entireWb ? _gridControl.Worksheets : new[] { _gridControl.CurrentWorksheet };
            int replacedCount = 0;

            foreach (var ws in sheets)
            {
                if (ws == null) continue;
                int maxR = Math.Min(ws.RowCount, 1000);
                int maxC = Math.Min(ws.ColumnCount, 100);

                for (int r = 0; r < maxR; r++)
                {
                    for (int c = 0; c < maxC; c++)
                    {
                        var cell = ws.Cells[r, c];
                        if (cell == null) continue;
                        string val = cell.Data?.ToString() ?? "";
                        if (string.IsNullOrEmpty(val)) continue;

                        if (matchEntire && string.Equals(val, query, comp))
                        {
                            if (double.TryParse(replaceWith, out double num))
                                cell.Data = num;
                            else
                                cell.Data = replaceWith;
                            replacedCount++;
                        }
                        else if (!matchEntire && val.IndexOf(query, comp) >= 0)
                        {
                            string newVal = val.Replace(query, replaceWith, comp);
                            if (double.TryParse(newVal, out double num))
                                cell.Data = num;
                            else
                                cell.Data = newVal;
                            replacedCount++;
                        }
                    }
                }
            }

            _gridControl.CurrentWorksheet?.Recalculate();
            TxtStatus.Text = $"All done. We made {replacedCount} replacement(s).";
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

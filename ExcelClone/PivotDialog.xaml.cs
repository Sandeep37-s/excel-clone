using System.Collections.Generic;
using System.Windows;
using Syncfusion.XlsIO;

namespace ExcelClone
{
    public partial class PivotDialog : Window
    {
        public string SourceRange => TxtSourceRange.Text.Trim();
        public int RowFieldIndex => CmbRowField.SelectedIndex;
        public int ColFieldIndex => CmbColField.SelectedIndex;
        public int ValueFieldIndex => CmbValueField.SelectedIndex;
        public bool IsNewSheet => RbNewSheet.IsChecked == true;
        public bool CreatePivotChart => ChkCreatePivotChart.IsChecked == true;

        public PivotSubtotalTypes Aggregation
        {
            get
            {
                return CmbAggregation.SelectedIndex switch
                {
                    1 => PivotSubtotalTypes.Count,
                    2 => PivotSubtotalTypes.Average,
                    3 => PivotSubtotalTypes.Max,
                    4 => PivotSubtotalTypes.Min,
                    _ => PivotSubtotalTypes.Sum
                };
            }
        }

        public PivotDialog(string defaultRange, List<string> columnHeaders)
        {
            InitializeComponent();
            TxtSourceRange.Text = defaultRange;

            foreach (var header in columnHeaders)
            {
                CmbRowField.Items.Add(header);
                CmbColField.Items.Add(header);
                CmbValueField.Items.Add(header);
            }

            if (columnHeaders.Count > 0) CmbRowField.SelectedIndex = 0;
            if (columnHeaders.Count > 1) CmbColField.SelectedIndex = columnHeaders.Count > 3 ? 3 : 1;
            if (columnHeaders.Count > 2) CmbValueField.SelectedIndex = columnHeaders.Count > 2 ? 2 : 0;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSourceRange.Text))
            {
                MessageBox.Show("Please specify a valid source range (e.g. A1:D9).", "Invalid Range", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}


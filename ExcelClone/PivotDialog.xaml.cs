using System.Collections.Generic;
using System.Windows;

namespace ExcelClone
{
    public enum PivotAggregationType
    {
        Sum,
        Count,
        Average,
        Max,
        Min
    }

    public partial class PivotDialog : Window
    {
        public string SourceRange => TxtSourceRange.Text.Trim();
        public int RowFieldIndex => CmbRowField.SelectedIndex;
        public int SecondaryRowFieldIndex => CmbSecondaryRowField.SelectedIndex - 1; // -1 means none
        public int ColFieldIndex => CmbColField.SelectedIndex;
        public int ValueFieldIndex => CmbValueField.SelectedIndex;
        public bool IsNewSheet => RbNewSheet.IsChecked == true;
        public bool CreatePivotChart => ChkCreatePivotChart.IsChecked == true;

        public PivotAggregationType Aggregation
        {
            get
            {
                return CmbAggregation.SelectedIndex switch
                {
                    1 => PivotAggregationType.Count,
                    2 => PivotAggregationType.Average,
                    3 => PivotAggregationType.Max,
                    4 => PivotAggregationType.Min,
                    _ => PivotAggregationType.Sum
                };
            }
        }

        public PivotDialog(string defaultRange, List<string> columnHeaders)
        {
            InitializeComponent();
            TxtSourceRange.Text = defaultRange;

            CmbSecondaryRowField.Items.Add("(None)");

            foreach (var header in columnHeaders)
            {
                CmbRowField.Items.Add(header);
                CmbSecondaryRowField.Items.Add(header);
                CmbColField.Items.Add(header);
                CmbValueField.Items.Add(header);
            }

            CmbSecondaryRowField.SelectedIndex = 0; // Default to None

            if (columnHeaders.Count > 0) CmbRowField.SelectedIndex = 0;
            if (columnHeaders.Count > 1) CmbColField.SelectedIndex = columnHeaders.Count > 3 ? 3 : 1;
            if (columnHeaders.Count > 2) CmbValueField.SelectedIndex = columnHeaders.Count > 2 ? 2 : 0;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtSourceRange.Text))
            {
                MessageBox.Show("Please select or enter a valid source range.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
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

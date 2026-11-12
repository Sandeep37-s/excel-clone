using System;
using System.Windows;
using Syncfusion.XlsIO;
using EFM_BOM_App.Services;

namespace EFM_BOM_App
{
    public partial class GoalSeekDialog : Window
    {
        private readonly IWorksheet _sheet;

        public GoalSeekDialog(IWorksheet sheet, string defaultTargetCell = "A1", string defaultChangingCell = "A2")
        {
            InitializeComponent();
            _sheet = sheet;
            TxtTargetCell.Text = defaultTargetCell;
            TxtChangingCell.Text = defaultChangingCell;
            TxtTargetValue.Text = "100";
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            string targetAddr = TxtTargetCell.Text.Trim();
            string changingAddr = TxtChangingCell.Text.Trim();

            if (string.IsNullOrEmpty(targetAddr) || string.IsNullOrEmpty(changingAddr))
            {
                MessageBox.Show("Please enter both target and changing cell references.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(TxtTargetValue.Text.Trim(), out double targetVal))
            {
                MessageBox.Show("Target value must be a valid number.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                TxtStatus.Text = "Solving Goal Seek...";
                bool success = NumericalAnalysisService.SolveGoalSeek(_sheet, targetAddr, targetVal, changingAddr, out double foundVal);

                if (success)
                {
                    TxtStatus.Text = $"Goal Seek converged!\nCell {changingAddr} set to: {foundVal:F4}";
                    DialogResult = true;
                    MessageBox.Show($"Goal Seek converged successfully!\n\nTarget cell {targetAddr} reached {targetVal}\nChanging cell {changingAddr} set to {foundVal:F6}", "Goal Seek Result", MessageBoxButton.OK, MessageBoxImage.Information);
                    Close();
                }
                else
                {
                    TxtStatus.Text = $"Goal Seek did not fully converge. Closest value: {foundVal:F4}";
                    MessageBox.Show($"Goal Seek could not find an exact solution within tolerance.\nClosest value found for {changingAddr}: {foundVal:F6}", "Goal Seek", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Goal Seek execution failed", ex);
                MessageBox.Show($"Goal Seek Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}


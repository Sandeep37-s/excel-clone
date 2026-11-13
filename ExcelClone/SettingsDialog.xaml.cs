using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using ExcelClone.Services;

namespace ExcelClone
{
    public partial class SettingsDialog : Window
    {
        public string SelectedTheme { get; private set; } = "Office2019Colorful";
        public int AutoSaveMinutes { get; private set; } = 5;
        public bool IsAutoCalc { get; private set; } = true;

        public SettingsDialog(string currentTheme, int autoSaveMins, bool isAutoCalc)
        {
            InitializeComponent();
            SelectedTheme = currentTheme;
            AutoSaveMinutes = autoSaveMins;
            IsAutoCalc = isAutoCalc;

            CmbTheme.SelectedIndex = currentTheme.Contains("DarkGray") ? 0 : 1;

            if (autoSaveMins == 1) CmbAutoSave.SelectedIndex = 0;
            else if (autoSaveMins == 5) CmbAutoSave.SelectedIndex = 1;
            else if (autoSaveMins == 10) CmbAutoSave.SelectedIndex = 2;
            else CmbAutoSave.SelectedIndex = 3;

            CmbCalcMode.SelectedIndex = isAutoCalc ? 0 : 1;

            try
            {
                string localKeyFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone", "syncfusion_license.txt");
                string baseKeyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "syncfusion_license.txt");
                if (File.Exists(baseKeyFile))
                {
                    TxtLicenseKey.Text = File.ReadAllText(baseKeyFile).Trim();
                }
                else if (File.Exists(localKeyFile))
                {
                    TxtLicenseKey.Text = File.ReadAllText(localKeyFile).Trim();
                }
            }
            catch { }
        }

        private void BtnOpenLogDir_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string logDir = Path.GetDirectoryName(LoggingService.GetLogFilePath()) ?? "";
                if (Directory.Exists(logDir))
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", logDir) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open directory: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnViewLogs_Click(object sender, RoutedEventArgs e)
        {
            string logs = LoggingService.ReadLogContents();
            var win = new Window
            {
                Title = "ExcelClone Event & Error Logs",
                Width = 700,
                Height = 450,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Content = new System.Windows.Controls.TextBox
                {
                    Text = logs,
                    IsReadOnly = true,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                    FontSize = 11,
                    VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                    Padding = new Thickness(10)
                }
            };
            win.ShowDialog();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SelectedTheme = CmbTheme.SelectedIndex == 0 ? "Office2019DarkGray" : "Office2019Colorful";

            switch (CmbAutoSave.SelectedIndex)
            {
                case 0: AutoSaveMinutes = 1; break;
                case 1: AutoSaveMinutes = 5; break;
                case 2: AutoSaveMinutes = 10; break;
                case 3: AutoSaveMinutes = 0; break; // Disabled
            }

            IsAutoCalc = (CmbCalcMode.SelectedIndex == 0);

            try
            {
                string key = TxtLicenseKey.Text.Trim();
                if (!string.IsNullOrWhiteSpace(key))
                {
                    string localDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone");
                    Directory.CreateDirectory(localDir);
                    File.WriteAllText(Path.Combine(localDir, "syncfusion_license.txt"), key);
                    File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "syncfusion_license.txt"), key);
                    Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(key);
                }
            }
            catch { }

            LoggingService.LogInfo($"Settings updated: Theme={SelectedTheme}, AutoSave={AutoSaveMinutes}m, AutoCalc={IsAutoCalc}");
            DialogResult = true;
            Close();
        }
    }
}


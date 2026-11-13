using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelClone.Services;

namespace ExcelClone
{
    public partial class SettingsDialog : Window
    {
        private UserSettings _settings;

        public SettingsDialog()
        {
            InitializeComponent();
            _settings = SettingsService.Load();
            LoadCurrentValues();
        }

        public SettingsDialog(string theme, int autoSaveMinutes, bool isAutoCalc) : this()
        {
        }

        private void LoadCurrentValues()
        {
            // Theme mapping
            CmbTheme.SelectedIndex = _settings.Theme.Contains("DarkGray") ? 0 : 1;

            // AutoSave mapping
            CmbAutoSave.SelectedIndex = _settings.AutoSaveMinutes switch
            {
                1 => 0,
                5 => 1,
                10 => 2,
                _ => 3
            };

            // Calculation mode
            CmbCalcMode.SelectedIndex = _settings.IsAutoCalc ? 0 : 1;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Save theme
            _settings.Theme = CmbTheme.SelectedIndex == 0 ? "Office2019DarkGray" : "Office2019Colorful";

            // Save autosave
            _settings.AutoSaveMinutes = CmbAutoSave.SelectedIndex switch
            {
                0 => 1,
                1 => 5,
                2 => 10,
                _ => 0
            };

            // Save calculation mode
            _settings.IsAutoCalc = CmbCalcMode.SelectedIndex == 0;

            SettingsService.Save(_settings);

            DialogResult = true;
            Close();
        }

        private void BtnOpenLogDir_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string logDir = LoggingService.LogDirectory;
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                Process.Start(new ProcessStartInfo
                {
                    FileName = logDir,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open log directory: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnViewLogs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string logDir = LoggingService.LogDirectory;
                string todayLog = Path.Combine(logDir, $"ExcelClone_{DateTime.Now:yyyyMMdd}.log");
                if (File.Exists(todayLog))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = todayLog,
                        UseShellExecute = true
                    });
                }
                else
                {
                    MessageBox.Show("No log file found for today.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open log file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

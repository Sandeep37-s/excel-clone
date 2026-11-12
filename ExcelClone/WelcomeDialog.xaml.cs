using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace ExcelClone
{
    public partial class WelcomeDialog : Window
    {
        public string? SelectedFileToOpen { get; private set; }
        public string? SelectedTemplate { get; private set; }
        public bool ShowOnStartup => ChkShowOnStartup.IsChecked == true;

        public class RecentItem
        {
            public string Name { get; set; } = "";
            public string DisplayPath { get; set; } = "";
            public string FullPath { get; set; } = "";
            public string DateModified { get; set; } = "";
        }

        private List<RecentItem> _allRecentItems = new();

        public WelcomeDialog(List<string> recentFiles, bool showOnStartup = true)
        {
            InitializeComponent();
            ChkShowOnStartup.IsChecked = showOnStartup;

            // Set greeting based on time of day
            int hour = DateTime.Now.Hour;
            if (hour < 12)
            {
                TxtGreeting.Text = "Good morning";
            }
            else if (hour < 17)
            {
                TxtGreeting.Text = "Good afternoon";
            }
            else
            {
                TxtGreeting.Text = "Good evening";
            }

            // Populate recent files
            PopulateRecentFiles(recentFiles);
        }

        private void PopulateRecentFiles(List<string> recentFiles)
        {
            _allRecentItems = new List<RecentItem>();

            // Add existing recent files from user history
            foreach (var path in recentFiles)
            {
                if (File.Exists(path))
                {
                    var fi = new FileInfo(path);
                    _allRecentItems.Add(new RecentItem
                    {
                        Name = Path.GetFileNameWithoutExtension(path),
                        DisplayPath = fi.Directory?.Name ?? Path.GetDirectoryName(path) ?? "",
                        FullPath = path,
                        DateModified = FormatDate(fi.LastWriteTime)
                    });
                }
            }

            // If recent list is sparse, add samples from the samples directory so the user gets a rich view
            string samplesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "samples");
            if (!Directory.Exists(samplesDir))
            {
                samplesDir = Path.Combine(Directory.GetCurrentDirectory(), "samples");
            }

            if (Directory.Exists(samplesDir))
            {
                var sampleFiles = Directory.GetFiles(samplesDir, "*.xlsx");
                foreach (var sample in sampleFiles)
                {
                    if (!_allRecentItems.Any(r => r.FullPath.Equals(sample, StringComparison.OrdinalIgnoreCase)))
                    {
                        var fi = new FileInfo(sample);
                        _allRecentItems.Add(new RecentItem
                        {
                            Name = Path.GetFileNameWithoutExtension(sample).Replace("_", " "),
                            DisplayPath = "Samples",
                            FullPath = sample,
                            DateModified = FormatDate(fi.LastWriteTime)
                        });
                    }
                }
            }

            ListRecentFiles.ItemsSource = _allRecentItems;
        }

        private string FormatDate(DateTime date)
        {
            if (date.Date == DateTime.Today)
            {
                return $"Today, {date:t}";
            }
            if (date.Date == DateTime.Today.AddDays(-1))
            {
                return "Yesterday";
            }
            return date.ToString("d MMMM");
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string q = TxtSearch.Text.Trim();
            TxtSearchPlaceholder.Visibility = string.IsNullOrEmpty(q) ? Visibility.Visible : Visibility.Collapsed;

            if (string.IsNullOrEmpty(q))
            {
                ListRecentFiles.ItemsSource = _allRecentItems;
            }
            else
            {
                ListRecentFiles.ItemsSource = _allRecentItems
                    .Where(r => r.Name.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                                r.DisplayPath.Contains(q, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }
        }

        private void TplBlank_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Blank";
            DialogResult = true;
            Close();
        }

        private void TplWelcomeTour_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Tour";
            DialogResult = true;
            Close();
        }

        private void TplFormulaTutorial_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Formula";
            DialogResult = true;
            Close();
        }

        private void TplFinance_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Finance";
            DialogResult = true;
            Close();
        }

        private void TplInventory_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Inventory";
            DialogResult = true;
            Close();
        }

        private void TplProject_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Project";
            DialogResult = true;
            Close();
        }

        private void ListRecentFiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ListRecentFiles.SelectedItem is RecentItem item)
            {
                SelectedFileToOpen = item.FullPath;
                DialogResult = true;
                Close();
            }
        }

        private void BtnNew_Click(object sender, RoutedEventArgs e)
        {
            SelectedTemplate = "Blank";
            DialogResult = true;
            Close();
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Workbooks (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|All Files (*.*)|*.*",
                Title = "Open Spreadsheet"
            };

            if (dlg.ShowDialog() == true)
            {
                SelectedFileToOpen = dlg.FileName;
                DialogResult = true;
                Close();
            }
        }

        private void BtnOptions_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SettingsDialog("Office2019Colorful", 5, true) { Owner = this };
            dlg.ShowDialog();
        }

        private void BtnContinue_Click(object sender, RoutedEventArgs e)
        {
            SelectedTemplate = "Blank";
            DialogResult = true;
            Close();
        }
    }
}

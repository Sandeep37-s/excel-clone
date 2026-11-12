using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;

namespace EFM_BOM_App
{
    public partial class WelcomeDialog : Window
    {
        public string? SelectedFileToOpen { get; private set; }
        public string? SelectedTemplate { get; private set; }
        public bool ShowOnStartup => ChkShowOnStartup.IsChecked == true;

        public record RecentItem(string Name, string Path);

        public WelcomeDialog(List<string> recentFiles, bool showOnStartup = true)
        {
            InitializeComponent();
            ChkShowOnStartup.IsChecked = showOnStartup;

            var items = recentFiles
                .Where(File.Exists)
                .Select(p => new RecentItem(System.IO.Path.GetFileName(p), p))
                .ToList();

            ListRecentFiles.ItemsSource = items;
        }

        private void TplBlank_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Blank";
            DialogResult = true;
            Close();
        }

        private void TplFinance_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "Finance";
            DialogResult = true;
            Close();
        }

        private void TplBOM_Click(object sender, MouseButtonEventArgs e)
        {
            SelectedTemplate = "BOM";
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
                SelectedFileToOpen = item.Path;
                DialogResult = true;
                Close();
            }
        }

        private void BtnBrowseOpen_Click(object sender, RoutedEventArgs e)
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

        private void BtnContinue_Click(object sender, RoutedEventArgs e)
        {
            SelectedTemplate = "Blank";
            DialogResult = true;
            Close();
        }
    }
}


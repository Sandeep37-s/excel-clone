using System;
using System.IO;
using System.Windows;
using Syncfusion.Licensing;
using Syncfusion.SfSkinManager;

namespace ExcelClone
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Register license key before any controls or windows initialize
            try
            {
                string? licenseKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
                string localAppDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone");
                string localKeyFile = Path.Combine(localAppDir, "syncfusion_license.txt");
                string baseKeyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "syncfusion_license.txt");

                if (string.IsNullOrWhiteSpace(licenseKey) && File.Exists(baseKeyFile))
                {
                    licenseKey = File.ReadAllText(baseKeyFile).Trim();
                }
                if (string.IsNullOrWhiteSpace(licenseKey) && File.Exists(localKeyFile))
                {
                    licenseKey = File.ReadAllText(localKeyFile).Trim();
                }

                if (!string.IsNullOrWhiteSpace(licenseKey))
                {
                    SyncfusionLicenseProvider.RegisterLicense(licenseKey);
                }
            }
            catch { }

            AppDomain.CurrentDomain.AssemblyResolve += (sender, resolveArgs) =>
            {
                if (resolveArgs.Name.StartsWith("System.Reflection.Metadata"))
                {
                    return typeof(System.Reflection.Metadata.MetadataReader).Assembly;
                }
                if (resolveArgs.Name.StartsWith("WindowsFormsIntegration"))
                {
                    string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WindowsFormsIntegration.dll");
                    if (File.Exists(path))
                    {
                        return System.Reflection.Assembly.LoadFrom(path);
                    }
                }
                return null;
            };

            // Apply visual style
            SfSkinManager.ApplyThemeAsDefaultStyle = true;

            base.OnStartup(e);
        }
    }
}

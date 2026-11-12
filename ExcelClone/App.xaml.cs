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

            // Register license key from environment variable or local config if available
            string? licenseKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
            if (string.IsNullOrWhiteSpace(licenseKey))
            {
                string keyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "syncfusion_license.txt");
                if (File.Exists(keyFile))
                {
                    licenseKey = File.ReadAllText(keyFile).Trim();
                }
            }

            if (!string.IsNullOrWhiteSpace(licenseKey))
            {
                SyncfusionLicenseProvider.RegisterLicense(licenseKey);
            }
        }
    }
}

using System;
using System.IO;
using System.Reflection;
using System.Windows;

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

            base.OnStartup(e);
        }
    }
}

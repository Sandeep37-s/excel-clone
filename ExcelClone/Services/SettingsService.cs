using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ExcelClone.Services
{
    public class UserSettings
    {
        public string Theme { get; set; } = "Office2019Colorful";
        public int AutoSaveMinutes { get; set; } = 5;
        public bool IsAutoCalc { get; set; } = true;
        public List<string> RecentFiles { get; set; } = new();
        public bool ShowWelcomeOnStartup { get; set; } = true;
    }

    public static class SettingsService
    {
        private static readonly string ConfigDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelClone");
        private static readonly string SettingsFile = Path.Combine(ConfigDir, "settings.json");

        public static UserSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    string json = File.ReadAllText(SettingsFile);
                    var settings = JsonSerializer.Deserialize<UserSettings>(json);
                    if (settings != null) return settings;
                }
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Failed to load user settings, reverting to defaults", ex);
            }
            return new UserSettings();
        }

        public static void Save(UserSettings settings)
        {
            try
            {
                Directory.CreateDirectory(ConfigDir);
                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFile, json);
                LoggingService.LogInfo("User settings persisted successfully.");
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Failed to save user settings", ex);
            }
        }
    }
}


using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace FileConverter
{
    public class UpdateChecker
    {
        private const string UpdateUrl = "https://your-website.com/updates/fileconverter.json";
        private readonly Version _currentVersion;

        public UpdateChecker(string currentVersion)
        {
            _currentVersion = new Version(currentVersion);
        }

        public async Task CheckForUpdatesAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var response = await client.GetStringAsync(UpdateUrl);
                    var updateInfo = JsonSerializer.Deserialize<UpdateInfo>(response);

                    var latestVersion = new Version(updateInfo.Version);
                    if (latestVersion > _currentVersion)
                    {
                        var result = MessageBox.Show(
                            $"A new version ({updateInfo.Version}) is available!\n\n{updateInfo.Description}\n\nWould you like to download it now?",
                            "Update Available",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Information);

                        if (result == MessageBoxResult.Yes)
                        {
                            // Open the download URL in the default browser
                            var processStartInfo = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = updateInfo.DownloadUrl,
                                UseShellExecute = true
                            };
                            System.Diagnostics.Process.Start(processStartInfo);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Silently fail - don't disturb user if update check fails
                Console.WriteLine($"Update check failed: {ex.Message}");
            }
        }

        public class UpdateInfo
        {
            public string Version { get; set; }
            public string DownloadUrl { get; set; }
            public string Description { get; set; }
        }
    }
}
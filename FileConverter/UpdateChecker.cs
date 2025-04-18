using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace FileConverter
{
    public class UpdateChecker
    {
        private readonly Version _currentVersion;
        private readonly string _repositoryOwner;
        private readonly string _repositoryName;
        private readonly HttpClient _client;

        public UpdateChecker(string currentVersion, string repositoryOwner, string repositoryName)
        {
            _currentVersion = new Version(currentVersion);
            _repositoryOwner = repositoryOwner;
            _repositoryName = repositoryName;

            _client = new HttpClient();
            // GitHub API requires a user agent
            _client.DefaultRequestHeaders.Add("User-Agent", "FileConverter-Update-Checker");
        }

        public async Task CheckForUpdatesAsync()
        {
            try
            {
                string apiUrl = $"https://api.github.com/repos/{_repositoryOwner}/{_repositoryName}/releases/latest";
                var response = await _client.GetStringAsync(apiUrl);

                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var releaseInfo = JsonSerializer.Deserialize<GitHubRelease>(response, options);

                if (releaseInfo == null || string.IsNullOrEmpty(releaseInfo.TagName))
                {
                    return;
                }

                // Remove 'v' prefix if present in tag name
                string versionString = releaseInfo.TagName.TrimStart('v');

                if (Version.TryParse(versionString, out Version latestVersion) &&
                    latestVersion > _currentVersion)
                {
                    var result = MessageBox.Show(
                        $"A new version ({releaseInfo.TagName}) is available!\n\n" +
                        $"{releaseInfo.Body ?? "No release notes available."}\n\n" +
                        "Would you like to download it now?",
                        "Update Available",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information);

                    if (result == MessageBoxResult.Yes)
                    {
                        // Find the asset that matches our installer
                        var installerAsset = releaseInfo.Assets?.FirstOrDefault(
                            asset => !string.IsNullOrEmpty(asset.Name) &&
                                   asset.Name.EndsWith(".exe") &&
                                   asset.Name.Contains("FileConverter"));

                        if (installerAsset != null && !string.IsNullOrEmpty(installerAsset.BrowserDownloadUrl))
                        {
                            // Open the download URL in the default browser
                            var processStartInfo = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = installerAsset.BrowserDownloadUrl,
                                UseShellExecute = true
                            };
                            System.Diagnostics.Process.Start(processStartInfo);
                        }
                        else if (!string.IsNullOrEmpty(releaseInfo.HtmlUrl))
                        {
                            // If no specific installer is found, open the release page
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = releaseInfo.HtmlUrl,
                                UseShellExecute = true
                            });
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

        // Models for GitHub API responses
        private class GitHubRelease
        {
            public string? TagName { get; set; }
            public string? Name { get; set; }
            public string? Body { get; set; }
            public string? HtmlUrl { get; set; }
            public GitHubAsset[]? Assets { get; set; }
        }

        private class GitHubAsset
        {
            public string? Name { get; set; }
            public string? BrowserDownloadUrl { get; set; }
        }
    }
}
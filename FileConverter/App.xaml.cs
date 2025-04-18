using System.Configuration;
using System.Data;
using System.Windows;
using System.Reflection;

namespace FileConverter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Get the current version from assembly
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            if (version != null)
            {
                string currentVersion = $"{version.Major}.{version.Minor}.{version.Build}";

                // Check for updates using GitHub
                var updateChecker = new UpdateChecker(
                    currentVersion,     // Current app version
                    "TSGCFO",          // Your GitHub username
                    "FileConverter"     // Your repository name
                );

                // Check for updates
                _ = updateChecker.CheckForUpdatesAsync();
            }
        }
    }
}
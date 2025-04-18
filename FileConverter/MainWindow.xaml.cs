// Bookmark: 1.0.0 (Import Statements - FileConverter Common Enums)
using FileConverter.Common.Enums; // Import file format enumeration for identifying input/output formats
// Bookmark: 1.0.1 (Import Statements - FileConverter Common Models)
using FileConverter.Common.Models; // Import data models for conversion parameters and results
// Bookmark: 1.0.2 (Import Statements - FileConverter Core Engine)
using FileConverter.Core.Engine; // Import conversion engine functionality
// Bookmark: 1.0.3 (Import Statements - Microsoft Win32)
using Microsoft.Win32; // Import Windows registry and dialog functionality
// Bookmark: 1.0.4 (Import Statements - System)
using System; // Import core C# functionality
// Bookmark: 1.0.5 (Import Statements - System Collections Generic)
using System.Collections.Generic; // Import generic collections like List, Dictionary
// Bookmark: 1.0.6 (Import Statements - System Configuration)
using System.Configuration; // Import configuration management
// Bookmark: 1.0.7 (Import Statements - System IO)
using System.IO; // Import file and directory operations
// Bookmark: 1.0.8 (Import Statements - System Linq)
using System.Linq; // Import LINQ query capabilities
// Bookmark: 1.0.9 (Import Statements - System Text)
using System.Text; // Import text handling capabilities
// Bookmark: 1.0.10 (Import Statements - System Threading)
using System.Threading; // Import threading capabilities for async operations
// Bookmark: 1.0.11 (Import Statements - System Windows)
using System.Windows; // Import WPF core functionality
// Bookmark: 1.0.12 (Import Statements - System Windows Controls)
using System.Windows.Controls; // Import WPF UI controls
// Bookmark: 1.0.13 (Import Statements - System Windows Media)
using System.Windows.Media; // Import WPF media and graphics
// Bookmark: 1.0.14 (Import Statements - System Xml Serialization)
using System.Xml.Serialization; // Import XML serialization for settings

// Bookmark: 1.1.0 (Namespace Declaration)
namespace FileConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml - This is the main application window
    /// that provides the user interface for file conversion operations.
    /// </summary>
    // Bookmark: 1.2.0 (MainWindow Class Declaration)
    public partial class MainWindow : Window
    {
        // Bookmark: 1.3.0 (Private Fields)
        private readonly ConversionEngine _conversionEngine; // Engine that performs the conversion operations
        private readonly Dictionary<string, FileFormat> _formatMapping; // Maps file extensions to format enum values
        private readonly Dictionary<FileFormat, string> _reverseFormatMapping; // Maps format enum values to display names
        private CancellationTokenSource? _cancellationTokenSource; // Used to cancel running conversions

        // Bookmark: 1.3.1 (Progress Tracking Fields)
        private DateTime _conversionStartTime; // Tracks when a conversion started for elapsed time calculation
        private System.Windows.Threading.DispatcherTimer? _estimationTimer; // Timer for updating time estimates

        // Bookmark: 1.3.2 (Settings Fields)
        private AppSettings _settings; // Application settings object
        private const string SettingsFilePath = "FileConverter.settings.xml"; // Path to the settings file

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// Sets up the UI components, initializes the conversion engine,
        /// and loads application settings.
        /// </summary>
        // Bookmark: 1.4.0 (Constructor)
        public MainWindow()
        {
            // Bookmark: 1.4.1 (Initialize UI Components)
            InitializeComponent(); // Initialize WPF UI components from XAML

            // Bookmark: 1.4.2 (Initialize Timer)
            _estimationTimer = new System.Windows.Threading.DispatcherTimer(); // Create timer for progress updates
            _estimationTimer.Interval = TimeSpan.FromSeconds(1); // Set timer to fire every second
            _estimationTimer.Tick += EstimationTimer_Tick; // Register event handler for timer tick events

            // Bookmark: 1.4.3 (Initialize Conversion Engine)
            _conversionEngine = ConversionEngineFactory.CreateWithDefaultConverters(); // Create conversion engine with default converters

            // Bookmark: 1.4.4 (Initialize Format Mappings)
            _formatMapping = new Dictionary<string, FileFormat>(StringComparer.OrdinalIgnoreCase) // Create case-insensitive dictionary
            {
                // Bookmark: 1.4.5 (Document Formats Mapping)
                // Map document format file extensions to enum values
                { ".pdf", FileFormat.Pdf },
                { ".doc", FileFormat.Doc },
                { ".docx", FileFormat.Docx },
                { ".rtf", FileFormat.Rtf },
                { ".odt", FileFormat.Odt },
                { ".txt", FileFormat.Txt },
                { ".html", FileFormat.Html },
                { ".htm", FileFormat.Html },
                { ".md", FileFormat.Md },
                
                // Bookmark: 1.4.6 (Spreadsheet Formats Mapping)
                // Map spreadsheet format file extensions to enum values
                { ".xlsx", FileFormat.Xlsx },
                { ".xls", FileFormat.Xls },
                { ".csv", FileFormat.Csv },
                { ".tsv", FileFormat.Tsv },
                
                // Bookmark: 1.4.7 (Data Formats Mapping)
                // Map data format file extensions to enum values
                { ".json", FileFormat.Json },
                { ".xml", FileFormat.Xml },
                { ".yaml", FileFormat.Yaml },
                { ".yml", FileFormat.Yaml },
                { ".ini", FileFormat.Ini },
                { ".toml", FileFormat.Toml }
            };

            // Bookmark: 1.4.8 (Create Reverse Format Mapping)
            // Create a reverse mapping from enum values to display names for UI
            _reverseFormatMapping = new Dictionary<FileFormat, string>();
            foreach (var kvp in _formatMapping)
            {
                if (!_reverseFormatMapping.ContainsKey(kvp.Value))
                {
                    _reverseFormatMapping.Add(kvp.Value, kvp.Key.TrimStart('.').ToUpper());
                }
            }

            // Bookmark: 1.4.9 (Initialize Settings)
            _settings = new AppSettings(); // Create default settings object

            // Bookmark: 1.4.10 (Load Application Settings)
            LoadSettings(); // Load saved settings from disk

            // Bookmark: 1.4.11 (Populate Format Combo Boxes)
            PopulateFormatComboBoxes(); // Fill input/output format dropdown lists

            // Bookmark: 1.4.12 (Add Menu Items)
            AddMenuItems(); // Add menu items to main menu

            // Bookmark: 1.4.13 (Apply Settings to UI)
            ApplySettings(); // Apply loaded settings to UI components
        }

        #region Settings Management

        /// <summary>
        /// Loads application settings from the settings file.
        /// If the file doesn't exist or there's an error, default settings are used.
        /// </summary>
        // Bookmark: 2.0.0 (Load Settings Method)
        private void LoadSettings()
        {
            try
            {
                // Bookmark: 2.0.1 (Check if Settings File Exists)
                if (File.Exists(SettingsFilePath)) // Check if settings file exists
                {
                    // Bookmark: 2.0.2 (Deserialize Settings)
                    using (var stream = new FileStream(SettingsFilePath, FileMode.Open)) // Open file stream
                    {
                        var serializer = new XmlSerializer(typeof(AppSettings)); // Create XML serializer for settings
                        var deserializedSettings = serializer.Deserialize(stream) as AppSettings; // Deserialize settings
                        if (deserializedSettings != null) // Check if deserialization was successful
                        {
                            _settings = deserializedSettings; // Apply deserialized settings
                        }
                        else
                        {
                            _settings = new AppSettings(); // Use default settings if deserialization failed
                            LogMessage("Deserialized settings were null. Using default settings."); // Log error
                        }
                    }
                    LogMessage("Settings loaded successfully."); // Log success
                }
                else
                {
                    // Bookmark: 2.0.3 (Use Default Settings)
                    _settings = new AppSettings(); // Create default settings
                    LogMessage("Using default settings."); // Log status
                }
            }
            catch (Exception ex)
            {
                // Bookmark: 2.0.4 (Handle Settings Load Error)
                _settings = new AppSettings(); // Use default settings on error
                LogMessage($"Error loading settings: {ex.Message}"); // Log error
            }
        }

        /// <summary>
        /// Saves current application settings to the settings file.
        /// </summary>
        // Bookmark: 2.1.0 (Save Settings Method)
        private void SaveSettings()
        {
            try
            {
                // Bookmark: 2.1.1 (Update Settings from UI)
                // Update settings from current UI state
                _settings.LastUsedInputDirectory = Path.GetDirectoryName(txtInputPath.Text) ?? "";
                _settings.LastUsedOutputDirectory = txtOutputPath.Tag?.ToString() ?? "";

                // Bookmark: 2.1.2 (Serialize Settings to File)
                using (var stream = new FileStream(SettingsFilePath, FileMode.Create)) // Create/overwrite settings file
                {
                    var serializer = new XmlSerializer(typeof(AppSettings)); // Create XML serializer
                    serializer.Serialize(stream, _settings); // Serialize settings to file
                }
                LogMessage("Settings saved successfully."); // Log success
            }
            catch (Exception ex)
            {
                // Bookmark: 2.1.3 (Handle Settings Save Error)
                LogMessage($"Error saving settings: {ex.Message}"); // Log error
            }
        }

        /// <summary>
        /// Applies loaded settings to the UI components.
        /// </summary>
        // Bookmark: 2.2.0 (Apply Settings Method)
        private void ApplySettings()
        {
            // Bookmark: 2.2.1 (Apply Input Directory)
            // Apply last used input directory if it exists
            if (!string.IsNullOrEmpty(_settings.LastUsedInputDirectory) &&
                Directory.Exists(_settings.LastUsedInputDirectory))
            {
                // We don't set the input path directly because it needs a file
                LogMessage($"Last used input directory: {_settings.LastUsedInputDirectory}");
            }

            // Bookmark: 2.2.2 (Apply Output Directory)
            // Apply last used output directory if it exists
            if (!string.IsNullOrEmpty(_settings.LastUsedOutputDirectory) &&
                Directory.Exists(_settings.LastUsedOutputDirectory))
            {
                txtOutputPath.Tag = _settings.LastUsedOutputDirectory; // Store directory in tag property
                LogMessage($"Last used output directory: {_settings.LastUsedOutputDirectory}");
            }

            // Bookmark: 2.2.3 (Apply UI Theme)
            // Apply dark mode theme settings (placeholder for future implementation)
            if (_settings.UseDarkMode)
            {
                // This would require implementing a theme system 
                // Apply dark mode theme here
            }

            // Bookmark: 2.2.4 (Apply Log Level)
            // Apply log level settings (placeholder for future implementation)
            // Implement different log levels if needed
        }

        /// <summary>
        /// Displays the settings dialog and applies changes if confirmed.
        /// </summary>
        // Bookmark: 2.3.0 (Show Settings Dialog Method)
        private void ShowSettingsDialog()
        {
            // Bookmark: 2.3.1 (Create and Show Settings Dialog)
            var settingsDialog = new SettingsDialog(_settings); // Create settings dialog with current settings
            if (settingsDialog.ShowDialog() == true) // Show dialog and check if user confirmed
            {
                // Bookmark: 2.3.2 (Apply Settings Changes)
                _settings = settingsDialog.Settings; // Update settings from dialog
                SaveSettings(); // Save to disk
                ApplySettings(); // Apply to UI
            }
        }

        /// <summary>
        /// Adds menu items to the main menu programmatically.
        /// </summary>
        // Bookmark: 2.4.0 (Add Menu Items Method)
        private void AddMenuItems()
        {
            // Bookmark: 2.4.1 (Check if Menu Exists)
            // Create a menu if it doesn't exist
            if (this.FindName("mainMenu") == null)
            {
                // Bookmark: 2.4.2 (Create Main Menu)
                var mainMenu = new Menu(); // Create menu control
                mainMenu.Name = "mainMenu"; // Set name for finding
                mainMenu.HorizontalAlignment = HorizontalAlignment.Left; // Align left
                mainMenu.VerticalAlignment = VerticalAlignment.Top; // Align top
                mainMenu.Height = 20; // Set height

                // Bookmark: 2.4.3 (Create File Menu)
                var fileMenu = new MenuItem { Header = "File" }; // Create File menu

                // Bookmark: 2.4.4 (Create Exit Menu Item)
                var exitMenuItem = new MenuItem { Header = "Exit" }; // Create Exit item
                exitMenuItem.Click += (s, e) => { this.Close(); }; // Set click handler to close app
                fileMenu.Items.Add(exitMenuItem); // Add to File menu

                // Bookmark: 2.4.5 (Create Settings Menu Item)
                var settingsMenuItem = new MenuItem { Header = "Settings" }; // Create Settings menu item
                settingsMenuItem.Click += (s, e) => { ShowSettingsDialog(); }; // Set click handler

                // Bookmark: 2.4.6 (Create Help Menu)
                var helpMenu = new MenuItem { Header = "Help" }; // Create Help menu

                // Bookmark: 2.4.7 (Create About Menu Item)
                var aboutMenuItem = new MenuItem { Header = "About" }; // Create About item
                aboutMenuItem.Click += (s, e) => { ShowAboutDialog(); }; // Set click handler
                helpMenu.Items.Add(aboutMenuItem); // Add to Help menu

                // Bookmark: 2.4.8 (Add All Menu Items)
                // Add menus to main menu
                mainMenu.Items.Add(fileMenu);
                mainMenu.Items.Add(settingsMenuItem);
                mainMenu.Items.Add(helpMenu);

                // Bookmark: 2.4.9 (Add Menu to Window)
                // Add menu to window
                if (this.Content is Grid mainGrid)
                {
                    // Bookmark: 2.4.10 (Insert Menu Row in Grid)
                    mainGrid.RowDefinitions.Insert(0, new RowDefinition { Height = new GridLength(20) }); // Add row at top
                    Grid.SetRow(mainMenu, 0); // Place menu in first row
                    mainGrid.Children.Add(mainMenu); // Add menu to grid

                    // Bookmark: 2.4.11 (Adjust Other Elements)
                    // Shift all other elements down by one row
                    foreach (UIElement element in mainGrid.Children)
                    {
                        if (element != mainMenu && Grid.GetRow(element) >= 0)
                        {
                            Grid.SetRow(element, Grid.GetRow(element) + 1); // Increment row
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Displays the About dialog with application information.
        /// </summary>
        // Bookmark: 2.5.0 (Show About Dialog Method)
        private void ShowAboutDialog()
        {
            // Bookmark: 2.5.1 (Display About Information)
            MessageBox.Show(
                "FileConverter\nVersion 1.0\n\nA versatile file format conversion utility.", // Message content
                "About FileConverter", // Dialog title
                MessageBoxButton.OK, // OK button only
                MessageBoxImage.Information); // Information icon
        }

        #endregion

        #region UI Event Handlers

        /// <summary>
        /// Handles the Copy Error button click by copying error details to clipboard.
        /// </summary>
        // Bookmark: 3.0.0 (Copy Error Button Click Handler)
        private void btnCopyError_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.0.1 (Check if Error Text Exists)
            if (!string.IsNullOrEmpty(txtErrorDetails.Text)) // Check if there's error text to copy
            {
                // Bookmark: 3.0.2 (Copy to Clipboard)
                Clipboard.SetText(txtErrorDetails.Text); // Copy to clipboard
                LogMessage("Error details copied to clipboard."); // Log action
            }
        }

        /// <summary>
        /// Handles the Close Error button click by hiding the error details panel.
        /// </summary>
        // Bookmark: 3.1.0 (Close Error Button Click Handler)
        private void btnCloseError_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.1.1 (Hide Error Panel)
            grpErrorDetails.Visibility = Visibility.Collapsed; // Hide error panel
        }

        /// <summary>
        /// Handles the Cancel button click by canceling the current conversion.
        /// </summary>
        // Bookmark: 3.2.0 (Cancel Button Click Handler)
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.2.1 (Check if Conversion Running)
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested) // Check if conversion is active
            {
                // Bookmark: 3.2.2 (Request Cancellation)
                _cancellationTokenSource.Cancel(); // Request cancellation
                LogMessage("Cancellation requested. Waiting for operation to stop..."); // Log action
                btnCancel.IsEnabled = false; // Disable cancel button
            }
        }

        /// <summary>
        /// Handles the Exit button click by closing the application.
        /// </summary>
        // Bookmark: 3.3.0 (Exit Button Click Handler)
        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.3.1 (Close Window)
            this.Close(); // Close the application window
        }

        /// <summary>
        /// Handles the Settings button click by showing the settings dialog.
        /// </summary>
        // Bookmark: 3.4.0 (Settings Button Click Handler)
        private void btnSettings_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.4.1 (Show Settings Dialog)
            ShowSettingsDialog(); // Show settings dialog
        }

        /// <summary>
        /// Handles the About button click by showing the about dialog.
        /// </summary>
        // Bookmark: 3.5.0 (About Button Click Handler)
        private void btnAbout_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.5.1 (Show About Dialog)
            ShowAboutDialog(); // Show about dialog
        }

        /// <summary>
        /// Handles timer ticks to update the time estimation display during conversion.
        /// </summary>
        // Bookmark: 3.6.0 (Estimation Timer Tick Handler)
        private void EstimationTimer_Tick(object? sender, EventArgs e)
        {
            // Bookmark: 3.6.1 (Check Progress and Calculate Remaining Time)
            if (progressBar.Value > 0) // Only estimate if progress has started
            {
                // Calculate estimated time remaining
                var elapsed = DateTime.Now - _conversionStartTime; // Time elapsed so far
                var estimatedTotal = elapsed.TotalSeconds / (progressBar.Value / 100.0); // Estimate total time based on progress
                var remaining = TimeSpan.FromSeconds(estimatedTotal - elapsed.TotalSeconds); // Calculate remaining time

                // Bookmark: 3.6.2 (Update Time Estimation Label)
                // Update the time estimation label if there's time remaining
                if (remaining.TotalSeconds > 0)
                {
                    txtTimeEstimation.Text = $"Estimated time remaining: {remaining:mm\\:ss}"; // Format as minutes:seconds
                }
            }
        }

        /// <summary>
        /// Handles the Convert button click by starting the conversion process.
        /// </summary>
        // Bookmark: 3.7.0 (Convert Button Click Handler)
        private async void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Bookmark: 3.7.1 (Validate Input and Output Paths)
                // Check if input path is provided
                if (string.IsNullOrWhiteSpace(txtInputPath.Text))
                {
                    MessageBox.Show("Please select an input file.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Check if output path is provided
                if (string.IsNullOrWhiteSpace(txtOutputPath.Text))
                {
                    MessageBox.Show("Please select an output file.", "Output Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Check if input file exists
                if (!File.Exists(txtInputPath.Text))
                {
                    MessageBox.Show("The input file does not exist.", "File Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Bookmark: 3.7.2 (Hide Previous Error Details)
                grpErrorDetails.Visibility = Visibility.Collapsed; // Hide error panel

                // Bookmark: 3.7.3 (Prepare UI for Conversion)
                btnConvert.IsEnabled = false; // Disable convert button
                btnCancel.IsEnabled = true; // Enable cancel button
                txtLog.Clear(); // Clear log
                progressBar.Value = 0; // Reset progress bar
                txtTimeEstimation.Visibility = Visibility.Visible; // Show time estimation

                // Bookmark: 3.7.4 (Initialize Conversion Tracking)
                _conversionStartTime = DateTime.Now; // Record start time

                // Bookmark: 3.7.5 (Start Estimation Timer)
                // Start the timer to update time remaining
                if (_estimationTimer != null)
                {
                    _estimationTimer.Start();
                }

                // Bookmark: 3.7.6 (Create Cancellation Token)
                _cancellationTokenSource = new CancellationTokenSource(); // Create new cancellation token

                // Bookmark: 3.7.7 (Create Progress Reporter)
                // Create progress reporter to update UI
                var progress = new Progress<ConversionProgress>(p =>
                {
                    // Update UI with progress
                    LogMessage($"{p.PercentComplete}%: {p.StatusMessage}");
                    progressBar.Value = p.PercentComplete;
                });

                // Bookmark: 3.7.8 (Get Conversion Parameters)
                var parameters = GetConversionParameters(); // Get parameters from UI

                // Bookmark: 3.7.9 (Perform Conversion)
                // Log start of conversion
                LogMessage($"Starting conversion: {txtInputPath.Text} -> {txtOutputPath.Text}");

                // Execute conversion
                var result = await _conversionEngine.ConvertFileAsync(
                    txtInputPath.Text,
                    txtOutputPath.Text,
                    parameters,
                    progress,
                    _cancellationTokenSource.Token);

                // Bookmark: 3.7.10 (Stop Estimation Timer)
                // Stop timer when conversion is done
                if (_estimationTimer != null)
                {
                    _estimationTimer.Stop();
                }
                txtTimeEstimation.Visibility = Visibility.Collapsed; // Hide time estimation

                // Bookmark: 3.7.11 (Handle Conversion Result)
                if (result.Success) // Check if conversion succeeded
                {
                    // Bookmark: 3.7.12 (Handle Successful Conversion)
                    LogMessage($"Conversion successful! Time: {result.ElapsedTime.TotalSeconds:F2} seconds");
                    MessageBox.Show(
                        $"Conversion completed successfully in {result.ElapsedTime.TotalSeconds:F2} seconds.\nOutput file: {Path.GetFullPath(txtOutputPath.Text)}",
                        "Conversion Complete",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    // Save the settings on successful conversion
                    SaveSettings();
                }
                else
                {
                    // Bookmark: 3.7.13 (Handle Failed Conversion)
                    LogMessage($"Conversion failed: {result.Error?.Message}");
                    ShowErrorDetails("Conversion failed", result.Error); // Show error details
                }
            }
            catch (Exception ex)
            {
                // Bookmark: 3.7.14 (Handle Unexpected Error)
                LogMessage($"Error: {ex.Message}");
                ShowErrorDetails("An unexpected error occurred", ex); // Show error details
            }
            finally
            {
                // Bookmark: 3.7.15 (Clean Up Conversion Resources)
                _cancellationTokenSource?.Dispose(); // Dispose token source
                _cancellationTokenSource = null; // Clear reference
                btnConvert.IsEnabled = true; // Re-enable convert button
                btnCancel.IsEnabled = false; // Disable cancel button
                if (_estimationTimer != null)
                {
                    _estimationTimer.Stop(); // Ensure timer is stopped
                }
                txtTimeEstimation.Visibility = Visibility.Collapsed; // Hide time estimation
            }
        }

        /// <summary>
        /// Handles the Browse Input button click by showing a file open dialog.
        /// </summary>
        // Bookmark: 3.8.0 (Browse Input Button Click Handler)
        private void btnBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.8.1 (Create Open File Dialog)
            var dialog = new OpenFileDialog
            {
                Title = "Select Input File", // Dialog title
                Filter = "All Files (*.*)|*.*", // File filter
                CheckFileExists = true // Ensure file exists
            };

            // Bookmark: 3.8.2 (Set Initial Directory from Settings)
            // Set initial directory from settings if available
            if (!string.IsNullOrEmpty(_settings.LastUsedInputDirectory) &&
                Directory.Exists(_settings.LastUsedInputDirectory))
            {
                dialog.InitialDirectory = _settings.LastUsedInputDirectory;
            }

            // Bookmark: 3.8.3 (Show Dialog and Process Result)
            if (dialog.ShowDialog() == true) // If user selected a file
            {
                // Bookmark: 3.8.4 (Set Input Path)
                txtInputPath.Text = dialog.FileName; // Set input path

                // Bookmark: 3.8.5 (Update Settings)
                // Update the last used input directory in settings
                _settings.LastUsedInputDirectory = Path.GetDirectoryName(dialog.FileName) ?? string.Empty;

                // Bookmark: 3.8.6 (Auto-Select Input Format)
                // Automatically select input format based on file extension
                string extension = Path.GetExtension(dialog.FileName).ToLowerInvariant();
                if (_formatMapping.TryGetValue(extension, out var format))
                {
                    string formatName = _reverseFormatMapping.GetValueOrDefault(format, format.ToString());
                    for (int i = 0; i < cmbInputFormat.Items.Count; i++)
                    {
                        if (cmbInputFormat.Items[i].ToString() == formatName)
                        {
                            cmbInputFormat.SelectedIndex = i;
                            break;
                        }
                    }
                }

                // Bookmark: 3.8.7 (Update Output Path)
                UpdateOutputPath(); // Update output path based on input
            }
        }

        /// <summary>
        /// Handles the Browse Output button click by showing a save dialog for the output directory.
        /// </summary>
        // Bookmark: 3.9.0 (Browse Output Button Click Handler)
        private void btnBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            // Bookmark: 3.9.1 (Create Save File Dialog)
            // Use standard SaveFileDialog but get just the directory
            var dialog = new SaveFileDialog
            {
                Title = "Select Output Directory", // Dialog title
                FileName = "Select This Directory", // Dummy filename
                Filter = "Directory|*.this.directory", // Filter to show it's for directory selection
                CheckPathExists = true // Ensure path exists
            };

            // Bookmark: 3.9.2 (Set Initial Directory from Settings)
            // Set initial directory from settings if available
            if (!string.IsNullOrEmpty(_settings.LastUsedOutputDirectory) &&
                Directory.Exists(_settings.LastUsedOutputDirectory))
            {
                dialog.InitialDirectory = _settings.LastUsedOutputDirectory;
            }

            // Bookmark: 3.9.3 (Show Dialog and Process Result)
            if (dialog.ShowDialog() == true) // If user selected a directory
            {
                // Bookmark: 3.9.4 (Extract Directory Path)
                // Extract just the directory path
                string? directoryPath = Path.GetDirectoryName(dialog.FileName);
                if (!string.IsNullOrEmpty(directoryPath))
                {
                    // Bookmark: 3.9.5 (Set Output Directory)
                    txtOutputPath.Tag = directoryPath; // Store directory in Tag property

                    // Bookmark: 3.9.6 (Update Settings)
                    // Update the last used output directory in settings
                    _settings.LastUsedOutputDirectory = directoryPath;

                    // Bookmark: 3.9.7 (Update Output Path)
                    UpdateOutputPath(); // Update output path based on directory
                }
            }
        }

        /// <summary>
        /// Handles the input format selection change by updating available output formats.
        /// </summary>
        // Bookmark: 3.10.0 (Input Format Selection Changed Handler)
        private void cmbInputFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Bookmark: 3.10.1 (Update Output Formats)
            UpdateOutputFormats(); // Update available output formats based on input format

            // Bookmark: 3.10.2 (Update Parameters Panel)
            UpdateParametersPanel(); // Update conversion parameters UI

            // Bookmark: 3.10.3 (Update Output Path)
            UpdateOutputPath(); // Update output path based on new format
        }

        /// <summary>
        /// Handles the output format selection change by updating parameters and output path.
        /// </summary>
        // Bookmark: 3.11.0 (Output Format Selection Changed Handler)
        private void cmbOutputFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Bookmark: 3.11.1 (Update Parameters Panel)
            UpdateParametersPanel(); // Update conversion parameters UI

            // Bookmark: 3.11.2 (Update Output Path)
            UpdateOutputPath(); // Update output path based on new format
        }

        #endregion

        #region UI Helper Methods

        /// <summary>
        /// Populates the input and output format combo boxes with available formats.
        /// </summary>
        // Bookmark: 4.0.0 (Populate Format Combo Boxes Method)
        private void PopulateFormatComboBoxes()
        {
            // Bookmark: 4.0.1 (Get Supported Conversion Paths)
            // Get all supported conversion paths from the engine
            var conversionPaths = _conversionEngine.GetSupportedConversionPaths().ToList();

            // Bookmark: 4.0.2 (Populate Input Formats)
            // Get distinct input formats and add to combo box
            var inputFormats = conversionPaths.Select(p => p.InputFormat).Distinct().ToList();
            foreach (var format in inputFormats.OrderBy(f => _reverseFormatMapping.GetValueOrDefault(f, f.ToString())))
            {
                cmbInputFormat.Items.Add(_reverseFormatMapping.GetValueOrDefault(format, format.ToString()));
            }

            // Bookmark: 4.0.3 (Disable Output Format Until Input Selected)
            // Initially no output formats are shown until an input format is selected
            cmbOutputFormat.IsEnabled = false;
        }

        /// <summary>
        /// Updates the output format combo box based on the selected input format.
        /// </summary>
        // Bookmark: 4.1.0 (Update Output Formats Method)
        private void UpdateOutputFormats()
        {
            // Bookmark: 4.1.1 (Clear Output Formats)
            cmbOutputFormat.Items.Clear(); // Clear existing items

            // Bookmark: 4.1.2 (Check Input Format Selected)
            // Get selected input format
            if (cmbInputFormat.SelectedItem == null)
                return;

            // Bookmark: 4.1.3 (Get Selected Format Name)
            var selectedFormatName = cmbInputFormat.SelectedItem.ToString();
            if (!_formatMapping.Values.Any(f => _reverseFormatMapping.GetValueOrDefault(f, f.ToString()) == selectedFormatName))
                return;

            // Bookmark: 4.1.4 (Get Format Enum from Name)
            var selectedFormat = _formatMapping.First(kvp =>
                _reverseFormatMapping.GetValueOrDefault(kvp.Value, kvp.Value.ToString()) == selectedFormatName).Value;

            // Bookmark: 4.1.5 (Get Available Output Formats)
            // Get all supported output formats for the selected input format
            var conversionPaths = _conversionEngine.GetSupportedConversionPaths().ToList();
            var outputFormats = conversionPaths
                .Where(p => p.InputFormat == selectedFormat)
                .Select(p => p.OutputFormat)
                .Distinct()
                .ToList();

            // Bookmark: 4.1.6 (Add Output Formats to ComboBox)
            foreach (var format in outputFormats.OrderBy(f => _reverseFormatMapping.GetValueOrDefault(f, f.ToString())))
            {
                cmbOutputFormat.Items.Add(_reverseFormatMapping.GetValueOrDefault(format, format.ToString()));
            }

            // Bookmark: 4.1.7 (Enable and Select First Output Format)
            cmbOutputFormat.IsEnabled = cmbOutputFormat.Items.Count > 0;
            if (cmbOutputFormat.Items.Count > 0)
                cmbOutputFormat.SelectedIndex = 0;
        }

        /// <summary>
        /// Updates the parameters panel based on the selected conversion type.
        /// </summary>
        // Bookmark: 4.2.0 (Update Parameters Panel Method)
        private void UpdateParametersPanel()
        {
            // Bookmark: 4.2.1 (Clear Parameters Panel)
            pnlParameters.Children.Clear(); // Clear existing parameters

            // Bookmark: 4.2.2 (Get Selected Formats)
            // Get current conversion path
            var inputFormat = GetSelectedInputFormat();
            var outputFormat = GetSelectedOutputFormat();

            // Bookmark: 4.2.3 (Check for Valid Formats)
            if (inputFormat == FileFormat.Unknown || outputFormat == FileFormat.Unknown)
                return;

            // Bookmark: 4.2.4 (Check if Conversion Supported)
            // Find the available converter for this path
            var conversionPaths = _conversionEngine.GetSupportedConversionPaths();
            var hasConverter = conversionPaths.Any(p =>
                p.InputFormat == inputFormat && p.OutputFormat == outputFormat);

            // Bookmark: 4.2.5 (Handle Unsupported Conversion)
            if (!hasConverter)
            {
                // Add a message that this conversion is not supported
                pnlParameters.Children.Add(new TextBlock
                {
                    Text = $"No converter available for {inputFormat} to {outputFormat} conversion.",
                    Foreground = Brushes.Red,
                    Margin = new Thickness(0, 5, 0, 5)
                });
                return;
            }

            // Bookmark: 4.2.6 (Add Parameters Based on Conversion Type)
            // Add parameters based on the selected conversion
            if (IsTextToHtmlSelected())
            {
                AddTextToHtmlParameters();
            }
            else if (IsTxtToTsvSelected())
            {
                AddTxtToTsvParameters();
            }
            else if (inputFormat == FileFormat.Txt && outputFormat == FileFormat.Csv)
            {
                AddTxtToCsvParameters();
            }
            else if (inputFormat == FileFormat.Csv && outputFormat == FileFormat.Tsv)
            {
                AddCsvToTsvParameters();
            }
            else if (inputFormat == FileFormat.Tsv && outputFormat == FileFormat.Csv)
            {
                AddTsvToCsvParameters();
            }
            else if (inputFormat == FileFormat.Json &&
                     (outputFormat == FileFormat.Csv || outputFormat == FileFormat.Tsv))
            {
                AddJsonToTabularParameters();
            }
            else if (inputFormat == FileFormat.Xml &&
                     (outputFormat == FileFormat.Csv || outputFormat == FileFormat.Tsv))
            {
                AddXmlToTabularParameters();
            }
            else if (inputFormat == FileFormat.Md && outputFormat == FileFormat.Tsv)
            {
                AddMarkdownToTsvParameters();
            }
            else if (inputFormat == FileFormat.Html && outputFormat == FileFormat.Tsv)
            {
                AddHtmlToTsvParameters();
            }
            else
            {
                // Bookmark: 4.2.7 (Add Default Message for Other Conversions)
                // Default message for other conversions
                pnlParameters.Children.Add(new TextBlock
                {
                    Text = "No additional parameters available for this conversion.",
                    Margin = new Thickness(0, 5, 0, 5)
                });
            }
        }

        #region Parameter Panel Creation Methods

        /// <summary>
        /// Adds a parameter control to the parameters panel with a label.
        /// </summary>
        // Bookmark: 4.3.0 (Add Parameter Method)
        private void AddParameter(string label, UIElement control, string tooltip = "")
        {
            // Bookmark: 4.3.1 (Create Parameter Panel)
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) }; // Create horizontal panel
            var labelBlock = new TextBlock { Text = label, Width = 120, VerticalAlignment = VerticalAlignment.Center }; // Create label

            // Bookmark: 4.3.2 (Add Tooltip if Provided)
            // Add tooltip if provided
            if (!string.IsNullOrEmpty(tooltip))
            {
                labelBlock.ToolTip = tooltip; // Set label tooltip
                control.SetValue(ToolTipProperty, tooltip); // Set control tooltip
            }

            // Bookmark: 4.3.3 (Add Controls to Panel)
            panel.Children.Add(labelBlock); // Add label to panel
            panel.Children.Add(control); // Add control to panel

            // Bookmark: 4.3.4 (Add Panel to Parameters Panel)
            pnlParameters.Children.Add(panel); // Add to main parameters panel
        }

        /// <summary>
        /// Adds a description text block to the parameters panel.
        /// </summary>
        // Bookmark: 4.4.0 (Add Description Text Method)
        private void AddDescriptionText(string text)
        {
            // Bookmark: 4.4.1 (Create Explanation Panel)
            var explanationPanel = new StackPanel { Margin = new Thickness(0, 5, 0, 5) }; // Create panel with margin

            // Bookmark: 4.4.2 (Add Text Block)
            explanationPanel.Children.Add(new TextBlock
            {
                Text = text, // Set text content
                TextWrapping = TextWrapping.Wrap // Enable text wrapping
            });

            // Bookmark: 4.4.3 (Add Panel to Parameters Panel)
            pnlParameters.Children.Add(explanationPanel); // Add to main parameters panel
        }

        /// <summary>
        /// Adds parameters for Text to HTML conversion.
        /// </summary>
        // Bookmark: 4.5.0 (Add Text to HTML Parameters Method)
        private void AddTextToHtmlParameters()
        {
            // Bookmark: 4.5.1 (Add Title Parameter)
            // Title parameter
            var titleBox = new TextBox { Name = "paramTitle", Width = 300, Text = "Converted Document" };
            AddParameter("Title:", titleBox, "The title to use in the HTML document");

            // Bookmark: 4.5.2 (Add Line Breaks Parameter)
            // Preserve line breaks parameter
            var lineBreaksCheckbox = new CheckBox { Name = "paramPreserveLineBreaks", IsChecked = true };
            AddParameter("Preserve Line Breaks:", lineBreaksCheckbox,
                "When checked, line breaks in the text file will be preserved in the HTML output");

            // Bookmark: 4.5.3 (Add CSS Parameter)
            // CSS parameter (advanced)
            var cssButton = new Button { Content = "Customize CSS...", Width = 150 };
            cssButton.Click += (s, e) => {
                // This would open a CSS editor dialog (not implemented)
                MessageBox.Show("CSS editor not implemented yet.", "Not Implemented",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            };
            AddParameter("Custom CSS:", cssButton, "Customize the CSS styling of the HTML document");
        }

        /// <summary>
        /// Adds parameters for Text to TSV conversion.
        /// </summary>
        // Bookmark: 4.6.0 (Add TXT to TSV Parameters Method)
        private void AddTxtToTsvParameters()
        {
            // Bookmark: 4.6.1 (Add Line Delimiter Parameter)
            // Line delimiter parameter
            var delimiterBox = new TextBox { Name = "paramLineDelimiter", Width = 300, Text = "" };
            AddParameter("Line Delimiter:", delimiterBox,
                "Character or string used to split each line into columns. Leave empty to treat each line as a single column.");

            // Bookmark: 4.6.2 (Add Explanation Text)
            // Explanation text
            AddDescriptionText("Leave the delimiter empty to treat each line as a single column. " +
                              "Enter a delimiter (like a comma) to split each line into multiple columns.");

            // Bookmark: 4.6.3 (Add Header Parameter)
            // First line as header parameter
            var headerCheckbox = new CheckBox { Name = "paramFirstLineHeader", IsChecked = false };
            AddParameter("First Line is Header:", headerCheckbox,
                "When checked, the first line of the text file will be treated as column headers");
        }

        /// <summary>
        /// Adds parameters for Text to CSV conversion.
        /// </summary>
        // Bookmark: 4.7.0 (Add TXT to CSV Parameters Method)
        private void AddTxtToCsvParameters()
        {
            // Bookmark: 4.7.1 (Add Line Delimiter Parameter)
            // Line delimiter parameter
            var delimiterBox = new TextBox { Name = "paramLineDelimiter", Width = 300, Text = "" };
            AddParameter("Line Delimiter:", delimiterBox,
                "Character or string used to split each line into columns. Leave empty to treat each line as a single column.");

            // Bookmark: 4.7.2 (Add Header Parameter)
            // First line as header parameter
            var headerCheckbox = new CheckBox { Name = "paramFirstLineHeader", IsChecked = false };
            AddParameter("First Line is Header:", headerCheckbox,
                "When checked, the first line of the text file will be treated as column headers");

            // Bookmark: 4.7.3 (Add CSV Delimiter Parameter)
            // CSV delimiter parameter
            var csvDelimiterBox = new TextBox { Name = "paramCsvDelimiter", Width = 300, Text = "," };
            AddParameter("CSV Delimiter:", csvDelimiterBox,
                "Character to use as delimiter in the CSV output (default is comma)");

            // Bookmark: 4.7.4 (Add Quote Character Parameter)
            // Quote character parameter
            var quoteBox = new TextBox { Name = "paramCsvQuote", Width = 300, Text = "\"" };
            AddParameter("CSV Quote Character:", quoteBox,
                "Character to use for quotes in the CSV output (default is double-quote)");
        }

        /// <summary>
        /// Adds parameters for CSV to TSV conversion.
        /// </summary>
        // Bookmark: 4.8.0 (Add CSV to TSV Parameters Method)
        private void AddCsvToTsvParameters()
        {
            // Bookmark: 4.8.1 (Add CSV Delimiter Parameter)
            // CSV delimiter parameter
            var csvDelimiterBox = new TextBox { Name = "paramCsvDelimiter", Width = 300, Text = "," };
            AddParameter("CSV Delimiter:", csvDelimiterBox,
                "Character used as delimiter in the input CSV file (default is comma)");

            // Bookmark: 4.8.2 (Add Quote Character Parameter)
            // Quote character parameter
            var quoteBox = new TextBox { Name = "paramCsvQuote", Width = 300, Text = "\"" };
            AddParameter("CSV Quote Character:", quoteBox,
                "Character used for quotes in the input CSV file (default is double-quote)");

            // Bookmark: 4.8.3 (Add Header Parameter)
            // Has header parameter
            var headerCheckbox = new CheckBox { Name = "paramHasHeader", IsChecked = true };
            AddParameter("Has Header Row:", headerCheckbox,
                "When checked, the first row of the CSV file will be treated as column headers");
        }

        /// <summary>
        /// Adds parameters for TSV to CSV conversion.
        /// </summary>
        // Bookmark: 4.9.0 (Add TSV to CSV Parameters Method)
        private void AddTsvToCsvParameters()
        {
            // Bookmark: 4.9.1 (Add CSV Delimiter Parameter)
            // CSV delimiter parameter
            var csvDelimiterBox = new TextBox { Name = "paramCsvDelimiter", Width = 300, Text = "," };
            AddParameter("CSV Delimiter:", csvDelimiterBox,
                "Character to use as delimiter in the CSV output (default is comma)");

            // Bookmark: 4.9.2 (Add Quote Character Parameter)
            // Quote character parameter
            var quoteBox = new TextBox { Name = "paramCsvQuote", Width = 300, Text = "\"" };
            AddParameter("CSV Quote Character:", quoteBox,
                "Character to use for quotes in the CSV output (default is double-quote)");

            // Bookmark: 4.9.3 (Add Header Parameter)
            // Has header parameter
            var headerCheckbox = new CheckBox { Name = "paramHasHeader", IsChecked = true };
            AddParameter("Has Header Row:", headerCheckbox,
                "When checked, the first row of the TSV file will be treated as column headers");
        }

        /// <summary>
        /// Adds parameters for JSON to tabular format conversion (CSV or TSV).
        /// </summary>
        // Bookmark: 4.10.0 (Add JSON to Tabular Parameters Method)
        private void AddJsonToTabularParameters()
        {
            // Bookmark: 4.10.1 (Add Array Path Parameter)
            // Array path parameter
            var arrayPathBox = new TextBox { Name = "paramArrayPath", Width = 300, Text = "" };
            AddParameter("Array Path:", arrayPathBox,
                "Path to the array in the JSON document (e.g., 'data.items'). Leave empty for auto-detection.");

            // Bookmark: 4.10.2 (Add Include Headers Parameter)
            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, column headers will be included in the output");

            // Bookmark: 4.10.3 (Add Max Nested Depth Parameter)
            // Max nested depth parameter
            var depthComboBox = new ComboBox { Name = "paramMaxDepth", Width = 300 };
            for (int i = 1; i <= 10; i++)
                depthComboBox.Items.Add(i);
            depthComboBox.SelectedIndex = 4; // Default to 5
            AddParameter("Max Nested Depth:", depthComboBox,
                "Maximum depth for nested objects. Higher values process deeper nested structures.");

            // Bookmark: 4.10.4 (Add Flatten Separator Parameter)
            // Flatten separator parameter
            var separatorBox = new TextBox { Name = "paramFlattenSeparator", Width = 300, Text = "." };
            AddParameter("Property Separator:", separatorBox,
                "Character used to separate nested property names in column headers");

            // Bookmark: 4.10.5 (Add Explanation Text)
            AddDescriptionText("For complex JSON structures, the converter will attempt to flatten nested objects " +
                              "into columns. The property separator defines how these nested paths appear in column headers.");
        }

        /// <summary>
        /// Adds parameters for XML to tabular format conversion (CSV or TSV).
        /// </summary>
        // Bookmark: 4.11.0 (Add XML to Tabular Parameters Method)
        private void AddXmlToTabularParameters()
        {
            // Bookmark: 4.11.1 (Add Root Element Path Parameter)
            // Root element path parameter
            var rootPathBox = new TextBox { Name = "paramRootElementPath", Width = 300, Text = "" };
            AddParameter("Element Path:", rootPathBox,
                "Path to the elements to convert (e.g., 'items/item'). Leave empty for auto-detection.");

            // Bookmark: 4.11.2 (Add Include Headers Parameter)
            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, column headers will be included in the output");

            // Bookmark: 4.11.3 (Add Explanation Text)
            AddDescriptionText("The converter will extract attributes and child elements from each XML element " +
                              "at the specified path and convert them to columns in the tabular format.");
        }

        /// <summary>
        /// Adds parameters for Markdown to TSV conversion.
        /// </summary>
        // Bookmark: 4.12.0 (Add Markdown to TSV Parameters Method)
        private void AddMarkdownToTsvParameters()
        {
            // Bookmark: 4.12.1 (Add Table Index Parameter)
            // Table index parameter
            var tableIndexBox = new ComboBox { Name = "paramTableIndex", Width = 300 };
            for (int i = 0; i < 10; i++)
                tableIndexBox.Items.Add(i);
            tableIndexBox.SelectedIndex = 0; // Default to first table
            AddParameter("Table Index:", tableIndexBox,
                "Index of the table to extract from the Markdown document (0 = first table)");

            // Bookmark: 4.12.2 (Add Include Headers Parameter)
            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, the header row of the Markdown table will be included in the output");

            // Bookmark: 4.12.3 (Add Explanation Text)
            AddDescriptionText("The converter will extract tables from Markdown files. If multiple tables exist, " +
                              "you can specify which one to convert using the Table Index parameter.");
        }

        /// <summary>
        /// Adds parameters for HTML to TSV conversion.
        /// </summary>
        // Bookmark: 4.13.0 (Add HTML to TSV Parameters Method)
        private void AddHtmlToTsvParameters()
        {
            // Bookmark: 4.13.1 (Add Table Index Parameter)
            // Table index parameter
            var tableIndexBox = new ComboBox { Name = "paramTableIndex", Width = 300 };
            for (int i = 0; i < 10; i++)
                tableIndexBox.Items.Add(i);
            tableIndexBox.SelectedIndex = 0; // Default to first table
            AddParameter("Table Index:", tableIndexBox,
                "Index of the table to extract from the HTML document (0 = first table)");

            // Bookmark: 4.13.2 (Add Include Headers Parameter)
            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, the header row of the HTML table will be included in the output");

            // Bookmark: 4.13.3 (Add Explanation Text)
            AddDescriptionText("The converter will extract tables from HTML files. If multiple tables exist, " +
                              "you can specify which one to convert using the Table Index parameter.");
        }

        #endregion

        /// <summary>
        /// Updates the output path based on the input file and output format.
        /// </summary>
        // Bookmark: 4.14.0 (Update Output Path Method)
        private void UpdateOutputPath()
        {
            // Bookmark: 4.14.1 (Check for Input File and Output Directory)
            // Only proceed if we have both an input file and output directory
            if (string.IsNullOrEmpty(txtInputPath.Text) || txtOutputPath.Tag == null)
            {
                return;
            }

            // Bookmark: 4.14.2 (Get Output Directory)
            string? outputDir = txtOutputPath.Tag?.ToString();
            if (string.IsNullOrEmpty(outputDir))
            {
                return;
            }

            // Bookmark: 4.14.3 (Get Input Filename Without Extension)
            // Get the input filename without extension
            string inputFilename = Path.GetFileNameWithoutExtension(txtInputPath.Text);

            // Bookmark: 4.14.4 (Get Output Format Extension)
            // Get the selected output format extension
            string outputExtension = ".txt"; // Default
            if (cmbOutputFormat.SelectedItem != null)
            {
                string? formatName = cmbOutputFormat.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(formatName))
                {
                    outputExtension = "." + formatName.ToLowerInvariant();
                }
            }

            // Bookmark: 4.14.5 (Create Output Path)
            // Combine to create the full output path
            string outputPath = Path.Combine(outputDir, inputFilename + outputExtension);

            // Bookmark: 4.14.6 (Update Output TextBox)
            // Update the text box
            txtOutputPath.Text = outputPath;
        }

        /// <summary>
        /// Checks if the Text to HTML conversion is selected.
        /// </summary>
        // Bookmark: 4.15.0 (Is Text to HTML Selected Method)
        private bool IsTextToHtmlSelected()
        {
            // Check if TXT to HTML conversion is selected
            return cmbInputFormat.SelectedItem?.ToString() == "TXT" &&
                   cmbOutputFormat.SelectedItem?.ToString() == "HTML";
        }

        /// <summary>
        /// Checks if the Text to TSV conversion is selected.
        /// </summary>
        // Bookmark: 4.16.0 (Is TXT to TSV Selected Method)
        private bool IsTxtToTsvSelected()
        {
            // Check if TXT to TSV conversion is selected
            return cmbInputFormat.SelectedItem?.ToString() == "TXT" &&
                   cmbOutputFormat.SelectedItem?.ToString() == "TSV";
        }

        /// <summary>
        /// Gets the currently selected input format as an enum value.
        /// </summary>
        // Bookmark: 4.17.0 (Get Selected Input Format Method)
        private FileFormat GetSelectedInputFormat()
        {
            // Bookmark: 4.17.1 (Check if Format is Selected)
            if (cmbInputFormat.SelectedItem == null)
                return FileFormat.Unknown;

            // Bookmark: 4.17.2 (Get Format Name from Selection)
            var formatName = cmbInputFormat.SelectedItem.ToString();

            // Bookmark: 4.17.3 (Find Format Enum from Name)
            foreach (var kvp in _formatMapping)
            {
                if (_reverseFormatMapping.TryGetValue(kvp.Value, out string? name) &&
                    name == formatName)
                {
                    return kvp.Value; // Return format enum
                }
            }

            // Bookmark: 4.17.4 (Return Unknown if Not Found)
            return FileFormat.Unknown;
        }

        /// <summary>
        /// Gets the currently selected output format as an enum value.
        /// </summary>
        // Bookmark: 4.18.0 (Get Selected Output Format Method)
        private FileFormat GetSelectedOutputFormat()
        {
            // Bookmark: 4.18.1 (Check if Format is Selected)
            if (cmbOutputFormat.SelectedItem == null)
                return FileFormat.Unknown;

            // Bookmark: 4.18.2 (Get Format Name from Selection)
            var formatName = cmbOutputFormat.SelectedItem.ToString();

            // Bookmark: 4.18.3 (Find Format Enum from Name)
            foreach (var kvp in _formatMapping)
            {
                if (_reverseFormatMapping.TryGetValue(kvp.Value, out string? name) &&
                    name == formatName)
                {
                    return kvp.Value; // Return format enum
                }
            }

            // Bookmark: 4.18.4 (Return Unknown if Not Found)
            return FileFormat.Unknown;
        }

        /// <summary>
        /// Finds a UI element by name in the visual tree.
        /// </summary>
        // Bookmark: 4.19.0 (Find Child Method)
        private T? FindChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            // Bookmark: 4.19.1 (Search Immediate Children)
            // Search immediate children
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                // Bookmark: 4.19.2 (Check if Child Matches)
                // Check if this is the child we're looking for
                if (child is FrameworkElement frameworkElement && frameworkElement.Name == childName)
                {
                    return child as T;
                }

                // Bookmark: 4.19.3 (Recursively Search Child's Children)
                // Recursively search this child's children
                var result = FindChild<T>(child, childName);
                if (result != null)
                    return result;
            }

            // Bookmark: 4.19.4 (Return Null if Not Found)
            return null;
        }

        /// <summary>
        /// Shows detailed error information in the error panel.
        /// </summary>
        // Bookmark: 4.20.0 (Show Error Details Method)
        private void ShowErrorDetails(string message, Exception? exception = null)
        {
            // Bookmark: 4.20.1 (Build Error Text)
            StringBuilder errorText = new StringBuilder();
            errorText.AppendLine($"Error: {message}");

            // Bookmark: 4.20.2 (Add Exception Details if Available)
            if (exception != null)
            {
                // Bookmark: 4.20.3 (Add Basic Exception Info)
                errorText.AppendLine();
                errorText.AppendLine("Technical Details:");
                errorText.AppendLine($"Type: {exception.GetType().FullName}");
                errorText.AppendLine($"Message: {exception.Message}");

                // Bookmark: 4.20.4 (Add Stack Trace if Available)
                if (exception.StackTrace != null)
                {
                    errorText.AppendLine();
                    errorText.AppendLine("Stack Trace:");
                    errorText.AppendLine(exception.StackTrace);
                }

                // Bookmark: 4.20.5 (Add Inner Exception Details if Available)
                // Include inner exception details if available
                var innerException = exception.InnerException;
                if (innerException != null)
                {
                    errorText.AppendLine();
                    errorText.AppendLine("Inner Exception:");
                    errorText.AppendLine($"Type: {innerException.GetType().FullName}");
                    errorText.AppendLine($"Message: {innerException.Message}");
                }
            }

            // Bookmark: 4.20.6 (Update UI with Error Details)
            txtErrorDetails.Text = errorText.ToString(); // Set error text
            grpErrorDetails.Visibility = Visibility.Visible; // Show error panel
        }

        /// <summary>
        /// Adds a message to the log with timestamp.
        /// </summary>
        // Bookmark: 4.21.0 (Log Message Method)
        private void LogMessage(string message)
        {
            // Bookmark: 4.21.1 (Add Timestamped Message to Log)
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n"); // Add message with timestamp
            txtLog.ScrollToEnd(); // Scroll to show latest message
        }

        #endregion

        /// <summary>
        /// Gets conversion parameters from the UI for the selected conversion type.
        /// </summary>
        // Bookmark: 5.0.0 (Get Conversion Parameters Method)
        private ConversionParameters GetConversionParameters()
        {
            // Bookmark: 5.0.1 (Create Parameters Object)
            var parameters = new ConversionParameters(); // Create empty parameters

            try
            {
                // Bookmark: 5.0.2 (Get Selected Formats)
                // Get current conversion path
                var inputFormat = GetSelectedInputFormat();
                var outputFormat = GetSelectedOutputFormat();

                // Bookmark: 5.0.3 (Check for Valid Formats)
                if (inputFormat == FileFormat.Unknown || outputFormat == FileFormat.Unknown)
                    return parameters;

                // Bookmark: 5.0.4 (Extract Appropriate Parameters)
                // Extract parameters based on the conversion type
                if (inputFormat == FileFormat.Txt && outputFormat == FileFormat.Html)
                {
                    ExtractTextToHtmlParameters(parameters);
                }
                else if (inputFormat == FileFormat.Txt && outputFormat == FileFormat.Tsv)
                {
                    ExtractTxtToTsvParameters(parameters);
                }
                else if (inputFormat == FileFormat.Txt && outputFormat == FileFormat.Csv)
                {
                    ExtractTxtToCsvParameters(parameters);
                }
                else if (inputFormat == FileFormat.Csv && outputFormat == FileFormat.Tsv)
                {
                    ExtractCsvToTsvParameters(parameters);
                }
                else if (inputFormat == FileFormat.Tsv && outputFormat == FileFormat.Csv)
                {
                    ExtractTsvToCsvParameters(parameters);
                }
                else if (inputFormat == FileFormat.Json &&
                         (outputFormat == FileFormat.Csv || outputFormat == FileFormat.Tsv))
                {
                    ExtractJsonToTabularParameters(parameters);
                }
                else if (inputFormat == FileFormat.Xml &&
                         (outputFormat == FileFormat.Csv || outputFormat == FileFormat.Tsv))
                {
                    ExtractXmlToTabularParameters(parameters);
                }
                else if (inputFormat == FileFormat.Md && outputFormat == FileFormat.Tsv)
                {
                    ExtractMarkdownToTsvParameters(parameters);
                }
                else if (inputFormat == FileFormat.Html && outputFormat == FileFormat.Tsv)
                {
                    ExtractHtmlToTsvParameters(parameters);
                }
            }
            catch (Exception ex)
            {
                // Bookmark: 5.0.5 (Handle Parameter Extraction Error)
                LogMessage($"Error extracting parameters: {ex.Message}"); // Log error
            }

            // Bookmark: 5.0.6 (Return Parameters)
            return parameters; // Return collected parameters
        }

        #region Parameter Extraction Methods

        /// <summary>
        /// Extracts parameters for Text to HTML conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.1.0 (Extract Text to HTML Parameters Method)
        private void ExtractTextToHtmlParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.1.1 (Extract Title Parameter)
            // Find the title parameter
            var titleBox = FindChild<TextBox>(pnlParameters, "paramTitle");
            if (titleBox != null)
            {
                parameters.AddParameter("title", titleBox.Text); // Add to parameters
            }

            // Bookmark: 5.1.2 (Extract Line Breaks Parameter)
            // Find the preserve line breaks parameter
            var lineBreaksCheckbox = FindChild<CheckBox>(pnlParameters, "paramPreserveLineBreaks");
            if (lineBreaksCheckbox != null)
            {
                parameters.AddParameter("preserveLineBreaks", lineBreaksCheckbox.IsChecked == true); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for Text to TSV conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.2.0 (Extract TXT to TSV Parameters Method)
        private void ExtractTxtToTsvParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.2.1 (Extract Line Delimiter Parameter)
            // Find the line delimiter parameter
            var delimiterBox = FindChild<TextBox>(pnlParameters, "paramLineDelimiter");
            if (delimiterBox != null)
            {
                parameters.AddParameter("lineDelimiter", delimiterBox.Text); // Add to parameters
            }

            // Bookmark: 5.2.2 (Extract Header Parameter)
            // Find the first line as header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramFirstLineHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("treatFirstLineAsHeader", headerCheckbox.IsChecked == true); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for Text to CSV conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.3.0 (Extract TXT to CSV Parameters Method)
        private void ExtractTxtToCsvParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.3.1 (Extract Line Delimiter Parameter)
            // Find the line delimiter parameter
            var delimiterBox = FindChild<TextBox>(pnlParameters, "paramLineDelimiter");
            if (delimiterBox != null)
            {
                parameters.AddParameter("lineDelimiter", delimiterBox.Text); // Add to parameters
            }

            // Bookmark: 5.3.2 (Extract Header Parameter)
            // Find the first line as header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramFirstLineHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("treatFirstLineAsHeader", headerCheckbox.IsChecked == true); // Add to parameters
            }

            // Bookmark: 5.3.3 (Extract CSV Delimiter Parameter)
            // Find the CSV delimiter parameter
            var csvDelimiterBox = FindChild<TextBox>(pnlParameters, "paramCsvDelimiter");
            if (csvDelimiterBox != null && !string.IsNullOrEmpty(csvDelimiterBox.Text))
            {
                parameters.AddParameter("csvDelimiter", csvDelimiterBox.Text[0]); // Add first character as delimiter
            }

            // Bookmark: 5.3.4 (Extract Quote Character Parameter)
            // Find the quote character parameter
            var quoteBox = FindChild<TextBox>(pnlParameters, "paramCsvQuote");
            if (quoteBox != null && !string.IsNullOrEmpty(quoteBox.Text))
            {
                parameters.AddParameter("csvQuote", quoteBox.Text[0]); // Add first character as quote
            }
        }

        /// <summary>
        /// Extracts parameters for CSV to TSV conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.4.0 (Extract CSV to TSV Parameters Method)
        private void ExtractCsvToTsvParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.4.1 (Extract CSV Delimiter Parameter)
            // Find the CSV delimiter parameter
            var csvDelimiterBox = FindChild<TextBox>(pnlParameters, "paramCsvDelimiter");
            if (csvDelimiterBox != null && !string.IsNullOrEmpty(csvDelimiterBox.Text))
            {
                parameters.AddParameter("csvDelimiter", csvDelimiterBox.Text[0]); // Add first character as delimiter
            }

            // Bookmark: 5.4.2 (Extract Quote Character Parameter)
            // Find the quote character parameter
            var quoteBox = FindChild<TextBox>(pnlParameters, "paramCsvQuote");
            if (quoteBox != null && !string.IsNullOrEmpty(quoteBox.Text))
            {
                parameters.AddParameter("csvQuote", quoteBox.Text[0]); // Add first character as quote
            }

            // Bookmark: 5.4.3 (Extract Header Parameter)
            // Find the has header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramHasHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("hasHeader", headerCheckbox.IsChecked == true); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for TSV to CSV conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.5.0 (Extract TSV to CSV Parameters Method)
        private void ExtractTsvToCsvParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.5.1 (Extract CSV Delimiter Parameter)
            // Find the CSV delimiter parameter
            var csvDelimiterBox = FindChild<TextBox>(pnlParameters, "paramCsvDelimiter");
            if (csvDelimiterBox != null && !string.IsNullOrEmpty(csvDelimiterBox.Text))
            {
                parameters.AddParameter("csvDelimiter", csvDelimiterBox.Text[0]); // Add first character as delimiter
            }

            // Bookmark: 5.5.2 (Extract Quote Character Parameter)
            // Find the quote character parameter
            var quoteBox = FindChild<TextBox>(pnlParameters, "paramCsvQuote");
            if (quoteBox != null && !string.IsNullOrEmpty(quoteBox.Text))
            {
                parameters.AddParameter("csvQuote", quoteBox.Text[0]); // Add first character as quote
            }

            // Bookmark: 5.5.3 (Extract Header Parameter)
            // Find the has header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramHasHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("hasHeader", headerCheckbox.IsChecked == true); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for JSON to tabular format conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.6.0 (Extract JSON to Tabular Parameters Method)
        private void ExtractJsonToTabularParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.6.1 (Extract Array Path Parameter)
            // Find the array path parameter
            var arrayPathBox = FindChild<TextBox>(pnlParameters, "paramArrayPath");
            if (arrayPathBox != null)
            {
                parameters.AddParameter("arrayPath", arrayPathBox.Text); // Add to parameters
            }

            // Bookmark: 5.6.2 (Extract Include Headers Parameter)
            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true); // Add to parameters
            }

            // Bookmark: 5.6.3 (Extract Max Depth Parameter)
            // Find the max depth parameter
            var depthComboBox = FindChild<ComboBox>(pnlParameters, "paramMaxDepth");
            if (depthComboBox != null && depthComboBox.SelectedItem != null)
            {
                parameters.AddParameter("maxDepth", (int)depthComboBox.SelectedItem); // Add to parameters
            }

            // Bookmark: 5.6.4 (Extract Separator Parameter)
            // Find the separator parameter
            var separatorBox = FindChild<TextBox>(pnlParameters, "paramFlattenSeparator");
            if (separatorBox != null)
            {
                parameters.AddParameter("flattenSeparator", separatorBox.Text); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for XML to tabular format conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.7.0 (Extract XML to Tabular Parameters Method)
        private void ExtractXmlToTabularParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.7.1 (Extract Root Element Path Parameter)
            // Find the root element path parameter
            var rootPathBox = FindChild<TextBox>(pnlParameters, "paramRootElementPath");
            if (rootPathBox != null)
            {
                parameters.AddParameter("rootElementPath", rootPathBox.Text); // Add to parameters
            }

            // Bookmark: 5.7.2 (Extract Include Headers Parameter)
            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for Markdown to TSV conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.8.0 (Extract Markdown to TSV Parameters Method)
        private void ExtractMarkdownToTsvParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.8.1 (Extract Table Index Parameter)
            // Find the table index parameter
            var tableIndexBox = FindChild<ComboBox>(pnlParameters, "paramTableIndex");
            if (tableIndexBox != null && tableIndexBox.SelectedItem != null)
            {
                parameters.AddParameter("tableIndex", (int)tableIndexBox.SelectedItem); // Add to parameters
            }

            // Bookmark: 5.8.2 (Extract Include Headers Parameter)
            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true); // Add to parameters
            }
        }

        /// <summary>
        /// Extracts parameters for HTML to TSV conversion from the UI.
        /// </summary>
        /// <param name="parameters">The parameters object to populate.</param>
        // Bookmark: 5.9.0 (Extract HTML to TSV Parameters Method)
        private void ExtractHtmlToTsvParameters(ConversionParameters parameters)
        {
            // Bookmark: 5.9.1 (Extract Table Index Parameter)
            // Find the table index parameter
            var tableIndexBox = FindChild<ComboBox>(pnlParameters, "paramTableIndex");
            if (tableIndexBox != null && tableIndexBox.SelectedItem != null)
            {
                parameters.AddParameter("tableIndex", (int)tableIndexBox.SelectedItem); // Add to parameters
            }

            // Bookmark: 5.9.2 (Extract Include Headers Parameter)
            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true); // Add to parameters
            }
        }

        #endregion
    }

    /// <summary>
    /// Class to store application settings that can be persisted to disk.
    /// Includes preferences for UI, directories, and conversion defaults.
    /// </summary>
    // Bookmark: 6.0.0 (App Settings Class Definition)
    [Serializable]
    public class AppSettings
    {
        // Bookmark: 6.0.1 (Last Used Input Directory Property)
        /// <summary>
        /// Gets or sets the last used input file directory.
        /// </summary>
        public string LastUsedInputDirectory { get; set; } = string.Empty;

        // Bookmark: 6.0.2 (Last Used Output Directory Property)
        /// <summary>
        /// Gets or sets the last used output file directory.
        /// </summary>
        public string LastUsedOutputDirectory { get; set; } = string.Empty;

        // Bookmark: 6.0.3 (Dark Mode Setting Property)
        /// <summary>
        /// Gets or sets a value indicating whether dark mode is enabled.
        /// </summary>
        public bool UseDarkMode { get; set; } = false;

        // Bookmark: 6.0.4 (Log Level Setting Property)
        /// <summary>
        /// Gets or sets the logging level.
        /// </summary>
        public string LogLevel { get; set; } = "Normal";

        // Bookmark: 6.0.5 (Remember Paths Setting Property)
        /// <summary>
        /// Gets or sets a value indicating whether to remember the last used paths.
        /// </summary>
        public bool RememberLastPaths { get; set; } = true;
    }

    /// <summary>
    /// Dialog for editing application settings.
    /// Provides UI elements for modifying settings values.
    /// </summary>
    // Bookmark: 7.0.0 (Settings Dialog Class Definition)
    public class SettingsDialog : Window
    {
        // Bookmark: 7.0.1 (Settings Property)
        /// <summary>
        /// Gets the current settings being edited.
        /// </summary>
        public AppSettings Settings { get; private set; }

        // Bookmark: 7.0.2 (UI Control Fields)
        private CheckBox? chkUseDarkMode; // Dark mode checkbox
        private ComboBox? cmbLogLevel; // Log level combo box
        private CheckBox? chkRememberPaths; // Remember paths checkbox

        /// <summary>
        /// Initializes a new instance of the SettingsDialog class.
        /// </summary>
        /// <param name="currentSettings">The current application settings to edit.</param>
        // Bookmark: 7.1.0 (Settings Dialog Constructor)
        public SettingsDialog(AppSettings currentSettings)
        {
            // Bookmark: 7.1.1 (Configure Dialog Window)
            Title = "Application Settings"; // Set window title
            Width = 450; // Set width
            Height = 300; // Set height
            WindowStartupLocation = WindowStartupLocation.CenterOwner; // Center on owner window

            // Bookmark: 7.1.2 (Clone Settings)
            // Clone settings to avoid changing originals if canceled
            Settings = new AppSettings
            {
                LastUsedInputDirectory = currentSettings.LastUsedInputDirectory,
                LastUsedOutputDirectory = currentSettings.LastUsedOutputDirectory,
                UseDarkMode = currentSettings.UseDarkMode,
                LogLevel = currentSettings.LogLevel,
                RememberLastPaths = currentSettings.RememberLastPaths
            };

            // Bookmark: 7.1.3 (Create UI Layout)
            // Create UI
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Bookmark: 7.1.4 (Create Settings Panel)
            var panel = new StackPanel { Margin = new Thickness(10) };

            // Bookmark: 7.1.5 (Create Dark Mode Setting)
            // Theme setting
            chkUseDarkMode = new CheckBox
            {
                Content = "Use Dark Mode", // Checkbox label
                IsChecked = Settings.UseDarkMode, // Set from settings
                Margin = new Thickness(0, 10, 0, 10) // Add margin
            };
            panel.Children.Add(chkUseDarkMode); // Add to panel

            // Bookmark: 7.1.6 (Create Log Level Setting)
            // Log level setting
            var logLevelPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 10) };
            logLevelPanel.Children.Add(new TextBlock { Text = "Log Level:", Width = 100, VerticalAlignment = VerticalAlignment.Center });

            cmbLogLevel = new ComboBox { Width = 200 };
            cmbLogLevel.Items.Add("Minimal");
            cmbLogLevel.Items.Add("Normal");
            cmbLogLevel.Items.Add("Verbose");
            cmbLogLevel.SelectedItem = Settings.LogLevel;
            logLevelPanel.Children.Add(cmbLogLevel);

            panel.Children.Add(logLevelPanel);

            // Bookmark: 7.1.7 (Create Remember Paths Setting)
            // Remember paths setting
            chkRememberPaths = new CheckBox
            {
                Content = "Remember Last Used Directories", // Checkbox label
                IsChecked = Settings.RememberLastPaths, // Set from settings
                Margin = new Thickness(0, 10, 0, 10) // Add margin
            };
            panel.Children.Add(chkRememberPaths); // Add to panel

            // Bookmark: 7.1.8 (Add Settings Panel to Grid)
            // Add settings panel to the grid
            grid.Children.Add(panel);
            Grid.SetRow(panel, 0);

            // Bookmark: 7.1.9 (Create Buttons Panel)
            // Create buttons
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10),
                Height = 30
            };

            // Bookmark: 7.1.10 (Create OK Button)
            var okButton = new Button
            {
                Content = "OK", // Button text
                Width = 80, // Set width
                Margin = new Thickness(5, 0, 5, 0), // Add margin
                IsDefault = true // Make default button (responds to Enter key)
            };
            okButton.Click += (s, e) =>
            {
                SaveSettings(); // Save settings changes
                DialogResult = true; // Close dialog with true result
            };

            // Bookmark: 7.1.11 (Create Cancel Button)
            var cancelButton = new Button
            {
                Content = "Cancel", // Button text
                Width = 80, // Set width
                Margin = new Thickness(5, 0, 5, 0), // Add margin
                IsCancel = true // Make cancel button (responds to Escape key)
            };
            cancelButton.Click += (s, e) => { DialogResult = false; }; // Close dialog with false result

            // Bookmark: 7.1.12 (Add Buttons to Panel)
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            // Bookmark: 7.1.13 (Add Buttons Panel to Grid)
            grid.Children.Add(buttonPanel);
            Grid.SetRow(buttonPanel, 1);

            // Bookmark: 7.1.14 (Set Dialog Content)
            // Set the content
            Content = grid;
        }

        /// <summary>
        /// Saves the current values from the UI controls to the Settings object.
        /// </summary>
        // Bookmark: 7.2.0 (Save Settings Method)
        private void SaveSettings()
        {
            // Bookmark: 7.2.1 (Save Dark Mode Setting)
            if (chkUseDarkMode != null)
                Settings.UseDarkMode = chkUseDarkMode.IsChecked == true;

            // Bookmark: 7.2.2 (Save Log Level Setting)
            if (cmbLogLevel != null && cmbLogLevel.SelectedItem != null)
                Settings.LogLevel = cmbLogLevel.SelectedItem.ToString() ?? string.Empty;

            // Bookmark: 7.2.3 (Save Remember Paths Setting)
            if (chkRememberPaths != null)
                Settings.RememberLastPaths = chkRememberPaths.IsChecked == true;
        }
    }
}
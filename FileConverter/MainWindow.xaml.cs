using FileConverter.Common.Enums;
using FileConverter.Common.Models;
using FileConverter.Core.Engine;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Serialization;

namespace FileConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ConversionEngine _conversionEngine;
        private readonly Dictionary<string, FileFormat> _formatMapping;
        private readonly Dictionary<FileFormat, string> _reverseFormatMapping;
        private CancellationTokenSource? _cancellationTokenSource;

        // New fields for improved progress reporting
        private DateTime _conversionStartTime;
        private System.Windows.Threading.DispatcherTimer? _estimationTimer;

        // Application settings
        private AppSettings _settings;
        private const string SettingsFilePath = "FileConverter.settings.xml";

        public MainWindow()
        {
            InitializeComponent();

            // Initialize the timer in constructor to fix CS8618
            _estimationTimer = new System.Windows.Threading.DispatcherTimer();
            _estimationTimer.Interval = TimeSpan.FromSeconds(1);
            _estimationTimer.Tick += EstimationTimer_Tick;

            // Create engine with default converters
            _conversionEngine = ConversionEngineFactory.CreateWithDefaultConverters();

            // Initialize format mappings
            _formatMapping = new Dictionary<string, FileFormat>(StringComparer.OrdinalIgnoreCase)
            {
                // Document formats
                { ".pdf", FileFormat.Pdf },
                { ".doc", FileFormat.Doc },
                { ".docx", FileFormat.Docx },
                { ".rtf", FileFormat.Rtf },
                { ".odt", FileFormat.Odt },
                { ".txt", FileFormat.Txt },
                { ".html", FileFormat.Html },
                { ".htm", FileFormat.Html },
                { ".md", FileFormat.Md },
                
                // Spreadsheet formats
                { ".xlsx", FileFormat.Xlsx },
                { ".xls", FileFormat.Xls },
                { ".csv", FileFormat.Csv },
                { ".tsv", FileFormat.Tsv },
                
                // Data formats
                { ".json", FileFormat.Json },
                { ".xml", FileFormat.Xml },
                { ".yaml", FileFormat.Yaml },
                { ".yml", FileFormat.Yaml },
                { ".ini", FileFormat.Ini },
                { ".toml", FileFormat.Toml }
            };

            // Create reverse mapping for displaying format names
            _reverseFormatMapping = new Dictionary<FileFormat, string>();
            foreach (var kvp in _formatMapping)
            {
                if (!_reverseFormatMapping.ContainsKey(kvp.Value))
                {
                    _reverseFormatMapping.Add(kvp.Value, kvp.Key.TrimStart('.').ToUpper());
                }
            }

            // Initialize _settings to avoid nullability issues
            _settings = new AppSettings();

            // Load application settings
            LoadSettings();

            // Populate format comboboxes
            PopulateFormatComboBoxes();

            // Add Settings and Help menu items
            AddMenuItems();

            // Apply settings to UI
            ApplySettings();
        }

        #region Settings Management

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    using (var stream = new FileStream(SettingsFilePath, FileMode.Open))
                    {
                        var serializer = new XmlSerializer(typeof(AppSettings));
                        var deserializedSettings = serializer.Deserialize(stream) as AppSettings;
                        if (deserializedSettings != null)
                        {
                            _settings = deserializedSettings;
                        }
                        else
                        {
                            _settings = new AppSettings();
                            LogMessage("Deserialized settings were null. Using default settings.");
                        }
                    }
                    LogMessage("Settings loaded successfully.");
                }
                else
                {
                    _settings = new AppSettings();
                    LogMessage("Using default settings.");
                }
            }
            catch (Exception ex)
            {
                _settings = new AppSettings();
                LogMessage($"Error loading settings: {ex.Message}");
            }
        }


        private void SaveSettings()
        {
            try
            {
                // Update settings from UI
                _settings.LastUsedInputDirectory = Path.GetDirectoryName(txtInputPath.Text) ?? "";
                _settings.LastUsedOutputDirectory = txtOutputPath.Tag?.ToString() ?? "";

                using (var stream = new FileStream(SettingsFilePath, FileMode.Create))
                {
                    var serializer = new XmlSerializer(typeof(AppSettings));
                    serializer.Serialize(stream, _settings);
                }
                LogMessage("Settings saved successfully.");
            }
            catch (Exception ex)
            {
                LogMessage($"Error saving settings: {ex.Message}");
            }
        }

        private void ApplySettings()
        {
            // Apply last used directories if they exist
            if (!string.IsNullOrEmpty(_settings.LastUsedInputDirectory) &&
                Directory.Exists(_settings.LastUsedInputDirectory))
            {
                // We don't set the input path directly because it needs a file
                LogMessage($"Last used input directory: {_settings.LastUsedInputDirectory}");
            }

            if (!string.IsNullOrEmpty(_settings.LastUsedOutputDirectory) &&
                Directory.Exists(_settings.LastUsedOutputDirectory))
            {
                txtOutputPath.Tag = _settings.LastUsedOutputDirectory;
                LogMessage($"Last used output directory: {_settings.LastUsedOutputDirectory}");
            }

            // Apply UI theme settings
            if (_settings.UseDarkMode)
            {
                // This would require implementing a theme system 
                // Apply dark mode theme here
            }

            // Apply log level
            // Implement different log levels if needed
        }

        private void ShowSettingsDialog()
        {
            var settingsDialog = new SettingsDialog(_settings);
            if (settingsDialog.ShowDialog() == true)
            {
                _settings = settingsDialog.Settings;
                SaveSettings();
                ApplySettings();
            }
        }

        private void AddMenuItems()
        {
            // Create a menu if it doesn't exist
            if (this.FindName("mainMenu") == null)
            {
                var mainMenu = new Menu();
                mainMenu.Name = "mainMenu";
                mainMenu.HorizontalAlignment = HorizontalAlignment.Left;
                mainMenu.VerticalAlignment = VerticalAlignment.Top;
                mainMenu.Height = 20;

                // Create File menu
                var fileMenu = new MenuItem { Header = "File" };

                var exitMenuItem = new MenuItem { Header = "Exit" };
                exitMenuItem.Click += (s, e) => { this.Close(); };
                fileMenu.Items.Add(exitMenuItem);

                // Create settings menu
                var settingsMenuItem = new MenuItem { Header = "Settings" };
                settingsMenuItem.Click += (s, e) => { ShowSettingsDialog(); };

                // Create Help menu
                var helpMenu = new MenuItem { Header = "Help" };

                var aboutMenuItem = new MenuItem { Header = "About" };
                aboutMenuItem.Click += (s, e) => { ShowAboutDialog(); };
                helpMenu.Items.Add(aboutMenuItem);

                // Add menus to main menu
                mainMenu.Items.Add(fileMenu);
                mainMenu.Items.Add(settingsMenuItem);
                mainMenu.Items.Add(helpMenu);

                // Add menu to window
                if (this.Content is Grid mainGrid)
                {
                    mainGrid.RowDefinitions.Insert(0, new RowDefinition { Height = new GridLength(20) });
                    Grid.SetRow(mainMenu, 0);
                    mainGrid.Children.Add(mainMenu);

                    // Adjust all other elements
                    foreach (UIElement element in mainGrid.Children)
                    {
                        if (element != mainMenu && Grid.GetRow(element) >= 0)
                        {
                            Grid.SetRow(element, Grid.GetRow(element) + 1);
                        }
                    }
                }
            }
        }

        private void ShowAboutDialog()
        {
            MessageBox.Show(
                "FileConverter\nVersion 1.0\n\nA versatile file format conversion utility.",
                "About FileConverter",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        #endregion

        #region UI Event Handlers

        private void btnCopyError_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtErrorDetails.Text))
            {
                Clipboard.SetText(txtErrorDetails.Text);
                LogMessage("Error details copied to clipboard.");
            }
        }

        private void btnCloseError_Click(object sender, RoutedEventArgs e)
        {
            grpErrorDetails.Visibility = Visibility.Collapsed;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
                LogMessage("Cancellation requested. Waiting for operation to stop...");
                btnCancel.IsEnabled = false;
            }
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnSettings_Click(object sender, RoutedEventArgs e)
        {
            ShowSettingsDialog();
        }

        private void btnAbout_Click(object sender, RoutedEventArgs e)
        {
            ShowAboutDialog();
        }

        // Fixed signature to use object? for sender parameter
        private void EstimationTimer_Tick(object? sender, EventArgs e)
        {
            if (progressBar.Value > 0)
            {
                // Calculate estimated time remaining
                var elapsed = DateTime.Now - _conversionStartTime;
                var estimatedTotal = elapsed.TotalSeconds / (progressBar.Value / 100.0);
                var remaining = TimeSpan.FromSeconds(estimatedTotal - elapsed.TotalSeconds);

                // Update the time estimation label
                if (remaining.TotalSeconds > 0)
                {
                    txtTimeEstimation.Text = $"Estimated time remaining: {remaining:mm\\:ss}";
                }
            }
        }

        private async void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Validate input and output paths
                if (string.IsNullOrWhiteSpace(txtInputPath.Text))
                {
                    MessageBox.Show("Please select an input file.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(txtOutputPath.Text))
                {
                    MessageBox.Show("Please select an output file.", "Output Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(txtInputPath.Text))
                {
                    MessageBox.Show("The input file does not exist.", "File Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Hide any previous error details
                grpErrorDetails.Visibility = Visibility.Collapsed;

                // Prepare for conversion
                btnConvert.IsEnabled = false;
                btnCancel.IsEnabled = true;
                txtLog.Clear();
                progressBar.Value = 0;
                txtTimeEstimation.Visibility = Visibility.Visible;

                // Record start time
                _conversionStartTime = DateTime.Now;

                // Use the timer that was already initialized in the constructor
                if (_estimationTimer != null)
                {
                    _estimationTimer.Start();
                }

                // Create cancellation token
                _cancellationTokenSource = new CancellationTokenSource();

                // Create progress reporter
                var progress = new Progress<ConversionProgress>(p =>
                {
                    // Update UI with progress
                    LogMessage($"{p.PercentComplete}%: {p.StatusMessage}");
                    progressBar.Value = p.PercentComplete;
                });

                // Get conversion parameters
                var parameters = GetConversionParameters();

                // Perform the conversion
                LogMessage($"Starting conversion: {txtInputPath.Text} -> {txtOutputPath.Text}");
                var result = await _conversionEngine.ConvertFileAsync(
                    txtInputPath.Text,
                    txtOutputPath.Text,
                    parameters,
                    progress,
                    _cancellationTokenSource.Token);

                // Stop timer
                if (_estimationTimer != null)
                {
                    _estimationTimer.Stop();
                }
                txtTimeEstimation.Visibility = Visibility.Collapsed;

                // Handle result
                if (result.Success)
                {
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
                    LogMessage($"Conversion failed: {result.Error?.Message}");
                    ShowErrorDetails("Conversion failed", result.Error);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error: {ex.Message}");
                ShowErrorDetails("An unexpected error occurred", ex);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                btnConvert.IsEnabled = true;
                btnCancel.IsEnabled = false;
                if (_estimationTimer != null)
                {
                    _estimationTimer.Stop();
                }
                txtTimeEstimation.Visibility = Visibility.Collapsed;
            }
        }

        private void btnBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Input File",
                Filter = "All Files (*.*)|*.*",
                CheckFileExists = true
            };

            // Set initial directory from settings if available
            if (!string.IsNullOrEmpty(_settings.LastUsedInputDirectory) &&
                Directory.Exists(_settings.LastUsedInputDirectory))
            {
                dialog.InitialDirectory = _settings.LastUsedInputDirectory;
            }

            if (dialog.ShowDialog() == true)
            {
                txtInputPath.Text = dialog.FileName;

                // Update the last used input directory in settings
                _settings.LastUsedInputDirectory = Path.GetDirectoryName(dialog.FileName) ?? string.Empty;

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

                // Update output path if we have an output directory
                UpdateOutputPath();
            }
        }

        private void btnBrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            // Use standard SaveFileDialog but get just the directory
            var dialog = new SaveFileDialog
            {
                Title = "Select Output Directory",
                FileName = "Select This Directory", // Dummy filename
                Filter = "Directory|*.this.directory",
                CheckPathExists = true
            };

            // Set initial directory from settings if available
            if (!string.IsNullOrEmpty(_settings.LastUsedOutputDirectory) &&
                Directory.Exists(_settings.LastUsedOutputDirectory))
            {
                dialog.InitialDirectory = _settings.LastUsedOutputDirectory;
            }

            if (dialog.ShowDialog() == true)
            {
                // Extract just the directory path
                string? directoryPath = Path.GetDirectoryName(dialog.FileName);
                if (!string.IsNullOrEmpty(directoryPath))
                {
                    txtOutputPath.Tag = directoryPath;

                    // Update the last used output directory in settings
                    _settings.LastUsedOutputDirectory = directoryPath;

                    UpdateOutputPath();
                }
            }
        }

        private void cmbInputFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateOutputFormats();
            UpdateParametersPanel();
            UpdateOutputPath();
        }

        private void cmbOutputFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateParametersPanel();
            UpdateOutputPath();
        }

        #endregion

        #region UI Helper Methods

        private void PopulateFormatComboBoxes()
        {
            // Get all supported conversion paths
            var conversionPaths = _conversionEngine.GetSupportedConversionPaths().ToList();

            // Get distinct input formats
            var inputFormats = conversionPaths.Select(p => p.InputFormat).Distinct().ToList();
            foreach (var format in inputFormats.OrderBy(f => _reverseFormatMapping.GetValueOrDefault(f, f.ToString())))
            {
                cmbInputFormat.Items.Add(_reverseFormatMapping.GetValueOrDefault(format, format.ToString()));
            }

            // Initially no output formats are shown until an input format is selected
            cmbOutputFormat.IsEnabled = false;
        }

        private void UpdateOutputFormats()
        {
            cmbOutputFormat.Items.Clear();

            // Get selected input format
            if (cmbInputFormat.SelectedItem == null)
                return;

            var selectedFormatName = cmbInputFormat.SelectedItem.ToString();
            if (!_formatMapping.Values.Any(f => _reverseFormatMapping.GetValueOrDefault(f, f.ToString()) == selectedFormatName))
                return;

            var selectedFormat = _formatMapping.First(kvp => _reverseFormatMapping.GetValueOrDefault(kvp.Value, kvp.Value.ToString()) == selectedFormatName).Value;

            // Get all supported output formats for the selected input format
            var conversionPaths = _conversionEngine.GetSupportedConversionPaths().ToList();
            var outputFormats = conversionPaths
                .Where(p => p.InputFormat == selectedFormat)
                .Select(p => p.OutputFormat)
                .Distinct()
                .ToList();

            foreach (var format in outputFormats.OrderBy(f => _reverseFormatMapping.GetValueOrDefault(f, f.ToString())))
            {
                cmbOutputFormat.Items.Add(_reverseFormatMapping.GetValueOrDefault(format, format.ToString()));
            }

            cmbOutputFormat.IsEnabled = cmbOutputFormat.Items.Count > 0;
            if (cmbOutputFormat.Items.Count > 0)
                cmbOutputFormat.SelectedIndex = 0;
        }

        private void UpdateParametersPanel()
        {
            pnlParameters.Children.Clear();

            // Get current conversion path
            var inputFormat = GetSelectedInputFormat();
            var outputFormat = GetSelectedOutputFormat();

            if (inputFormat == FileFormat.Unknown || outputFormat == FileFormat.Unknown)
                return;

            // Find the available converter for this path
            var conversionPaths = _conversionEngine.GetSupportedConversionPaths();
            var hasConverter = conversionPaths.Any(p =>
                p.InputFormat == inputFormat && p.OutputFormat == outputFormat);

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
                // Default message for other conversions
                pnlParameters.Children.Add(new TextBlock
                {
                    Text = "No additional parameters available for this conversion.",
                    Margin = new Thickness(0, 5, 0, 5)
                });
            }
        }

        #region Parameter Panel Creation Methods

        private void AddParameter(string label, UIElement control, string tooltip = "")
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };
            var labelBlock = new TextBlock { Text = label, Width = 120, VerticalAlignment = VerticalAlignment.Center };

            // Add tooltip if provided
            if (!string.IsNullOrEmpty(tooltip))
            {
                labelBlock.ToolTip = tooltip;
                control.SetValue(ToolTipProperty, tooltip);
            }

            panel.Children.Add(labelBlock);
            panel.Children.Add(control);
            pnlParameters.Children.Add(panel);
        }

        private void AddDescriptionText(string text)
        {
            var explanationPanel = new StackPanel { Margin = new Thickness(0, 5, 0, 5) };
            explanationPanel.Children.Add(new TextBlock
            {
                Text = text,
                TextWrapping = TextWrapping.Wrap
            });
            pnlParameters.Children.Add(explanationPanel);
        }

        private void AddTextToHtmlParameters()
        {
            // Title parameter
            var titleBox = new TextBox { Name = "paramTitle", Width = 300, Text = "Converted Document" };
            AddParameter("Title:", titleBox, "The title to use in the HTML document");

            // Preserve line breaks parameter
            var lineBreaksCheckbox = new CheckBox { Name = "paramPreserveLineBreaks", IsChecked = true };
            AddParameter("Preserve Line Breaks:", lineBreaksCheckbox,
                "When checked, line breaks in the text file will be preserved in the HTML output");

            // CSS parameter (advanced)
            var cssButton = new Button { Content = "Customize CSS...", Width = 150 };
            cssButton.Click += (s, e) => {
                // This would open a CSS editor dialog
                MessageBox.Show("CSS editor not implemented yet.", "Not Implemented",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            };
            AddParameter("Custom CSS:", cssButton, "Customize the CSS styling of the HTML document");
        }

        private void AddTxtToTsvParameters()
        {
            // Line delimiter parameter
            var delimiterBox = new TextBox { Name = "paramLineDelimiter", Width = 300, Text = "" };
            AddParameter("Line Delimiter:", delimiterBox,
                "Character or string used to split each line into columns. Leave empty to treat each line as a single column.");

            // Explanation text
            AddDescriptionText("Leave the delimiter empty to treat each line as a single column. " +
                              "Enter a delimiter (like a comma) to split each line into multiple columns.");

            // First line as header parameter
            var headerCheckbox = new CheckBox { Name = "paramFirstLineHeader", IsChecked = false };
            AddParameter("First Line is Header:", headerCheckbox,
                "When checked, the first line of the text file will be treated as column headers");
        }

        private void AddTxtToCsvParameters()
        {
            // Line delimiter parameter
            var delimiterBox = new TextBox { Name = "paramLineDelimiter", Width = 300, Text = "" };
            AddParameter("Line Delimiter:", delimiterBox,
                "Character or string used to split each line into columns. Leave empty to treat each line as a single column.");

            // First line as header parameter
            var headerCheckbox = new CheckBox { Name = "paramFirstLineHeader", IsChecked = false };
            AddParameter("First Line is Header:", headerCheckbox,
                "When checked, the first line of the text file will be treated as column headers");

            // CSV delimiter parameter
            var csvDelimiterBox = new TextBox { Name = "paramCsvDelimiter", Width = 300, Text = "," };
            AddParameter("CSV Delimiter:", csvDelimiterBox,
                "Character to use as delimiter in the CSV output (default is comma)");

            // Quote character parameter
            var quoteBox = new TextBox { Name = "paramCsvQuote", Width = 300, Text = "\"" };
            AddParameter("CSV Quote Character:", quoteBox,
                "Character to use for quotes in the CSV output (default is double-quote)");
        }

        private void AddCsvToTsvParameters()
        {
            // CSV delimiter parameter
            var csvDelimiterBox = new TextBox { Name = "paramCsvDelimiter", Width = 300, Text = "," };
            AddParameter("CSV Delimiter:", csvDelimiterBox,
                "Character used as delimiter in the input CSV file (default is comma)");

            // Quote character parameter
            var quoteBox = new TextBox { Name = "paramCsvQuote", Width = 300, Text = "\"" };
            AddParameter("CSV Quote Character:", quoteBox,
                "Character used for quotes in the input CSV file (default is double-quote)");

            // Has header parameter
            var headerCheckbox = new CheckBox { Name = "paramHasHeader", IsChecked = true };
            AddParameter("Has Header Row:", headerCheckbox,
                "When checked, the first row of the CSV file will be treated as column headers");
        }

        private void AddTsvToCsvParameters()
        {
            // CSV delimiter parameter
            var csvDelimiterBox = new TextBox { Name = "paramCsvDelimiter", Width = 300, Text = "," };
            AddParameter("CSV Delimiter:", csvDelimiterBox,
                "Character to use as delimiter in the CSV output (default is comma)");

            // Quote character parameter
            var quoteBox = new TextBox { Name = "paramCsvQuote", Width = 300, Text = "\"" };
            AddParameter("CSV Quote Character:", quoteBox,
                "Character to use for quotes in the CSV output (default is double-quote)");

            // Has header parameter
            var headerCheckbox = new CheckBox { Name = "paramHasHeader", IsChecked = true };
            AddParameter("Has Header Row:", headerCheckbox,
                "When checked, the first row of the TSV file will be treated as column headers");
        }

        private void AddJsonToTabularParameters()
        {
            // Array path parameter
            var arrayPathBox = new TextBox { Name = "paramArrayPath", Width = 300, Text = "" };
            AddParameter("Array Path:", arrayPathBox,
                "Path to the array in the JSON document (e.g., 'data.items'). Leave empty for auto-detection.");

            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, column headers will be included in the output");

            // Max nested depth parameter
            var depthComboBox = new ComboBox { Name = "paramMaxDepth", Width = 300 };
            for (int i = 1; i <= 10; i++)
                depthComboBox.Items.Add(i);
            depthComboBox.SelectedIndex = 4; // Default to 5
            AddParameter("Max Nested Depth:", depthComboBox,
                "Maximum depth for nested objects. Higher values process deeper nested structures.");

            // Flatten separator parameter
            var separatorBox = new TextBox { Name = "paramFlattenSeparator", Width = 300, Text = "." };
            AddParameter("Property Separator:", separatorBox,
                "Character used to separate nested property names in column headers");

            AddDescriptionText("For complex JSON structures, the converter will attempt to flatten nested objects " +
                              "into columns. The property separator defines how these nested paths appear in column headers.");
        }

        private void AddXmlToTabularParameters()
        {
            // Root element path parameter
            var rootPathBox = new TextBox { Name = "paramRootElementPath", Width = 300, Text = "" };
            AddParameter("Element Path:", rootPathBox,
                "Path to the elements to convert (e.g., 'items/item'). Leave empty for auto-detection.");

            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, column headers will be included in the output");

            AddDescriptionText("The converter will extract attributes and child elements from each XML element " +
                              "at the specified path and convert them to columns in the tabular format.");
        }

        private void AddMarkdownToTsvParameters()
        {
            // Table index parameter
            var tableIndexBox = new ComboBox { Name = "paramTableIndex", Width = 300 };
            for (int i = 0; i < 10; i++)
                tableIndexBox.Items.Add(i);
            tableIndexBox.SelectedIndex = 0; // Default to first table
            AddParameter("Table Index:", tableIndexBox,
                "Index of the table to extract from the Markdown document (0 = first table)");

            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, the header row of the Markdown table will be included in the output");

            AddDescriptionText("The converter will extract tables from Markdown files. If multiple tables exist, " +
                              "you can specify which one to convert using the Table Index parameter.");
        }

        private void AddHtmlToTsvParameters()
        {
            // Table index parameter
            var tableIndexBox = new ComboBox { Name = "paramTableIndex", Width = 300 };
            for (int i = 0; i < 10; i++)
                tableIndexBox.Items.Add(i);
            tableIndexBox.SelectedIndex = 0; // Default to first table
            AddParameter("Table Index:", tableIndexBox,
                "Index of the table to extract from the HTML document (0 = first table)");

            // Include headers parameter
            var headerCheckbox = new CheckBox { Name = "paramIncludeHeaders", IsChecked = true };
            AddParameter("Include Headers:", headerCheckbox,
                "When checked, the header row of the HTML table will be included in the output");

            AddDescriptionText("The converter will extract tables from HTML files. If multiple tables exist, " +
                              "you can specify which one to convert using the Table Index parameter.");
        }

        #endregion

        private void UpdateOutputPath()
        {
            // Only proceed if we have both an input file and output directory
            if (string.IsNullOrEmpty(txtInputPath.Text) || txtOutputPath.Tag == null)
            {
                return;
            }

            string? outputDir = txtOutputPath.Tag?.ToString();
            if (string.IsNullOrEmpty(outputDir))
            {
                return;
            }

            // Get the input filename without extension
            string inputFilename = Path.GetFileNameWithoutExtension(txtInputPath.Text);

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

            // Combine to create the full output path
            string outputPath = Path.Combine(outputDir, inputFilename + outputExtension);

            // Update the text box
            txtOutputPath.Text = outputPath;
        }

        private bool IsTextToHtmlSelected()
        {
            return cmbInputFormat.SelectedItem?.ToString() == "TXT" &&
                   cmbOutputFormat.SelectedItem?.ToString() == "HTML";
        }

        private bool IsTxtToTsvSelected()
        {
            return cmbInputFormat.SelectedItem?.ToString() == "TXT" &&
                   cmbOutputFormat.SelectedItem?.ToString() == "TSV";
        }

        private FileFormat GetSelectedInputFormat()
        {
            if (cmbInputFormat.SelectedItem == null)
                return FileFormat.Unknown;

            var formatName = cmbInputFormat.SelectedItem.ToString();
            foreach (var kvp in _formatMapping)
            {
                if (_reverseFormatMapping.TryGetValue(kvp.Value, out string? name) &&
                    name == formatName)
                {
                    return kvp.Value;
                }
            }

            return FileFormat.Unknown;
        }

        private FileFormat GetSelectedOutputFormat()
        {
            if (cmbOutputFormat.SelectedItem == null)
                return FileFormat.Unknown;

            var formatName = cmbOutputFormat.SelectedItem.ToString();
            foreach (var kvp in _formatMapping)
            {
                if (_reverseFormatMapping.TryGetValue(kvp.Value, out string? name) &&
                    name == formatName)
                {
                    return kvp.Value;
                }
            }

            return FileFormat.Unknown;
        }

        private T? FindChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            // Search immediate children
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                // Check if this is the child we're looking for
                if (child is FrameworkElement frameworkElement && frameworkElement.Name == childName)
                {
                    return child as T;
                }

                // Recursively search this child's children
                var result = FindChild<T>(child, childName);
                if (result != null)
                    return result;
            }

            return null;
        }

        private void ShowErrorDetails(string message, Exception? exception = null)
        {
            StringBuilder errorText = new StringBuilder();
            errorText.AppendLine($"Error: {message}");

            if (exception != null)
            {
                errorText.AppendLine();
                errorText.AppendLine("Technical Details:");
                errorText.AppendLine($"Type: {exception.GetType().FullName}");
                errorText.AppendLine($"Message: {exception.Message}");

                if (exception.StackTrace != null)
                {
                    errorText.AppendLine();
                    errorText.AppendLine("Stack Trace:");
                    errorText.AppendLine(exception.StackTrace);
                }

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

            txtErrorDetails.Text = errorText.ToString();
            grpErrorDetails.Visibility = Visibility.Visible;
        }

        private void LogMessage(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            txtLog.ScrollToEnd();
        }

        #endregion

        private ConversionParameters GetConversionParameters()
        {
            var parameters = new ConversionParameters();

            try
            {
                // Get current conversion path
                var inputFormat = GetSelectedInputFormat();
                var outputFormat = GetSelectedOutputFormat();

                if (inputFormat == FileFormat.Unknown || outputFormat == FileFormat.Unknown)
                    return parameters;

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
                LogMessage($"Error extracting parameters: {ex.Message}");
            }

            return parameters;
        }

        #region Parameter Extraction Methods

        private void ExtractTextToHtmlParameters(ConversionParameters parameters)
        {
            // Find the title parameter
            var titleBox = FindChild<TextBox>(pnlParameters, "paramTitle");
            if (titleBox != null)
            {
                parameters.AddParameter("title", titleBox.Text);
            }

            // Find the preserve line breaks parameter
            var lineBreaksCheckbox = FindChild<CheckBox>(pnlParameters, "paramPreserveLineBreaks");
            if (lineBreaksCheckbox != null)
            {
                parameters.AddParameter("preserveLineBreaks", lineBreaksCheckbox.IsChecked == true);
            }
        }

        private void ExtractTxtToTsvParameters(ConversionParameters parameters)
        {
            // Find the line delimiter parameter
            var delimiterBox = FindChild<TextBox>(pnlParameters, "paramLineDelimiter");
            if (delimiterBox != null)
            {
                parameters.AddParameter("lineDelimiter", delimiterBox.Text);
            }

            // Find the first line as header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramFirstLineHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("treatFirstLineAsHeader", headerCheckbox.IsChecked == true);
            }
        }

        private void ExtractTxtToCsvParameters(ConversionParameters parameters)
        {
            // Find the line delimiter parameter
            var delimiterBox = FindChild<TextBox>(pnlParameters, "paramLineDelimiter");
            if (delimiterBox != null)
            {
                parameters.AddParameter("lineDelimiter", delimiterBox.Text);
            }

            // Find the first line as header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramFirstLineHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("treatFirstLineAsHeader", headerCheckbox.IsChecked == true);
            }

            // Find the CSV delimiter parameter
            var csvDelimiterBox = FindChild<TextBox>(pnlParameters, "paramCsvDelimiter");
            if (csvDelimiterBox != null && !string.IsNullOrEmpty(csvDelimiterBox.Text))
            {
                parameters.AddParameter("csvDelimiter", csvDelimiterBox.Text[0]);
            }

            // Find the quote character parameter
            var quoteBox = FindChild<TextBox>(pnlParameters, "paramCsvQuote");
            if (quoteBox != null && !string.IsNullOrEmpty(quoteBox.Text))
            {
                parameters.AddParameter("csvQuote", quoteBox.Text[0]);
            }
        }

        private void ExtractCsvToTsvParameters(ConversionParameters parameters)
        {
            // Find the CSV delimiter parameter
            var csvDelimiterBox = FindChild<TextBox>(pnlParameters, "paramCsvDelimiter");
            if (csvDelimiterBox != null && !string.IsNullOrEmpty(csvDelimiterBox.Text))
            {
                parameters.AddParameter("csvDelimiter", csvDelimiterBox.Text[0]);
            }

            // Find the quote character parameter
            var quoteBox = FindChild<TextBox>(pnlParameters, "paramCsvQuote");
            if (quoteBox != null && !string.IsNullOrEmpty(quoteBox.Text))
            {
                parameters.AddParameter("csvQuote", quoteBox.Text[0]);
            }

            // Find the has header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramHasHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("hasHeader", headerCheckbox.IsChecked == true);
            }
        }

        private void ExtractTsvToCsvParameters(ConversionParameters parameters)
        {
            // Find the CSV delimiter parameter
            var csvDelimiterBox = FindChild<TextBox>(pnlParameters, "paramCsvDelimiter");
            if (csvDelimiterBox != null && !string.IsNullOrEmpty(csvDelimiterBox.Text))
            {
                parameters.AddParameter("csvDelimiter", csvDelimiterBox.Text[0]);
            }

            // Find the quote character parameter
            var quoteBox = FindChild<TextBox>(pnlParameters, "paramCsvQuote");
            if (quoteBox != null && !string.IsNullOrEmpty(quoteBox.Text))
            {
                parameters.AddParameter("csvQuote", quoteBox.Text[0]);
            }

            // Find the has header parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramHasHeader");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("hasHeader", headerCheckbox.IsChecked == true);
            }
        }

        private void ExtractJsonToTabularParameters(ConversionParameters parameters)
        {
            // Find the array path parameter
            var arrayPathBox = FindChild<TextBox>(pnlParameters, "paramArrayPath");
            if (arrayPathBox != null)
            {
                parameters.AddParameter("arrayPath", arrayPathBox.Text);
            }

            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true);
            }

            // Find the max depth parameter
            var depthComboBox = FindChild<ComboBox>(pnlParameters, "paramMaxDepth");
            if (depthComboBox != null && depthComboBox.SelectedItem != null)
            {
                parameters.AddParameter("maxDepth", (int)depthComboBox.SelectedItem);
            }

            // Find the separator parameter
            var separatorBox = FindChild<TextBox>(pnlParameters, "paramFlattenSeparator");
            if (separatorBox != null)
            {
                parameters.AddParameter("flattenSeparator", separatorBox.Text);
            }
        }

        private void ExtractXmlToTabularParameters(ConversionParameters parameters)
        {
            // Find the root element path parameter
            var rootPathBox = FindChild<TextBox>(pnlParameters, "paramRootElementPath");
            if (rootPathBox != null)
            {
                parameters.AddParameter("rootElementPath", rootPathBox.Text);
            }

            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true);
            }
        }

        private void ExtractMarkdownToTsvParameters(ConversionParameters parameters)
        {
            // Find the table index parameter
            var tableIndexBox = FindChild<ComboBox>(pnlParameters, "paramTableIndex");
            if (tableIndexBox != null && tableIndexBox.SelectedItem != null)
            {
                parameters.AddParameter("tableIndex", (int)tableIndexBox.SelectedItem);
            }

            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true);
            }
        }

        private void ExtractHtmlToTsvParameters(ConversionParameters parameters)
        {
            // Find the table index parameter
            var tableIndexBox = FindChild<ComboBox>(pnlParameters, "paramTableIndex");
            if (tableIndexBox != null && tableIndexBox.SelectedItem != null)
            {
                parameters.AddParameter("tableIndex", (int)tableIndexBox.SelectedItem);
            }

            // Find the include headers parameter
            var headerCheckbox = FindChild<CheckBox>(pnlParameters, "paramIncludeHeaders");
            if (headerCheckbox != null)
            {
                parameters.AddParameter("includeHeaders", headerCheckbox.IsChecked == true);
            }
        }

        #endregion
    }

    /// <summary>
    /// Class to store application settings
    /// </summary>
    [Serializable]
    public class AppSettings
    {
        public string LastUsedInputDirectory { get; set; } = string.Empty;
        public string LastUsedOutputDirectory { get; set; } = string.Empty;
        public bool UseDarkMode { get; set; } = false;
        public string LogLevel { get; set; } = "Normal";
        public bool RememberLastPaths { get; set; } = true;
    }

    /// <summary>
    /// Dialog for editing application settings
    /// </summary>
    public class SettingsDialog : Window
    {
        public AppSettings Settings { get; private set; }

        private CheckBox? chkUseDarkMode;
        private ComboBox? cmbLogLevel;
        private CheckBox? chkRememberPaths;

        public SettingsDialog(AppSettings currentSettings)
        {
            Title = "Application Settings";
            Width = 450;
            Height = 300;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            // Clone settings to avoid changing originals if canceled
            Settings = new AppSettings
            {
                LastUsedInputDirectory = currentSettings.LastUsedInputDirectory,
                LastUsedOutputDirectory = currentSettings.LastUsedOutputDirectory,
                UseDarkMode = currentSettings.UseDarkMode,
                LogLevel = currentSettings.LogLevel,
                RememberLastPaths = currentSettings.RememberLastPaths
            };

            // Create UI
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var panel = new StackPanel { Margin = new Thickness(10) };

            // Theme setting
            chkUseDarkMode = new CheckBox
            {
                Content = "Use Dark Mode",
                IsChecked = Settings.UseDarkMode,
                Margin = new Thickness(0, 10, 0, 10)
            };
            panel.Children.Add(chkUseDarkMode);

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

            // Remember paths setting
            chkRememberPaths = new CheckBox
            {
                Content = "Remember Last Used Directories",
                IsChecked = Settings.RememberLastPaths,
                Margin = new Thickness(0, 10, 0, 10)
            };
            panel.Children.Add(chkRememberPaths);

            // Add settings panel to the grid
            grid.Children.Add(panel);
            Grid.SetRow(panel, 0);

            // Create buttons
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10),
                Height = 30
            };

            var okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Margin = new Thickness(5, 0, 5, 0),
                IsDefault = true
            };
            okButton.Click += (s, e) =>
            {
                SaveSettings();
                DialogResult = true;
            };

            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 80,
                Margin = new Thickness(5, 0, 5, 0),
                IsCancel = true
            };
            cancelButton.Click += (s, e) => { DialogResult = false; };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            grid.Children.Add(buttonPanel);
            Grid.SetRow(buttonPanel, 1);

            // Set the content
            Content = grid;
        }

        private void SaveSettings()
        {
            if (chkUseDarkMode != null)
                Settings.UseDarkMode = chkUseDarkMode.IsChecked == true;

            if (cmbLogLevel != null && cmbLogLevel.SelectedItem != null)
                Settings.LogLevel = cmbLogLevel.SelectedItem.ToString() ?? string.Empty;

            if (chkRememberPaths != null)
                Settings.RememberLastPaths = chkRememberPaths.IsChecked == true;
        }
    }
}
using FileConverter.Common.Enums;
using FileConverter.Common.Models;
using FileConverter.Core.Engine;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

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

        public MainWindow()
        {
            InitializeComponent();

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

            // Populate format comboboxes
            PopulateFormatComboBoxes();
        }

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

            // Add parameters based on the selected conversion
            if (IsTextToHtmlSelected())
            {
                // Title parameter
                var titlePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };
                titlePanel.Children.Add(new TextBlock { Text = "Title:", Width = 120, VerticalAlignment = VerticalAlignment.Center });
                titlePanel.Children.Add(new TextBox { Name = "paramTitle", Width = 300, Text = "Converted Document" });
                pnlParameters.Children.Add(titlePanel);

                // Preserve line breaks parameter
                var lineBreaksPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };
                lineBreaksPanel.Children.Add(new TextBlock { Text = "Preserve Line Breaks:", Width = 120, VerticalAlignment = VerticalAlignment.Center });
                lineBreaksPanel.Children.Add(new CheckBox { Name = "paramPreserveLineBreaks", IsChecked = true });
                pnlParameters.Children.Add(lineBreaksPanel);
            }
            // Add parameters for TxtToTsv conversion
            else if (IsTxtToTsvSelected())
            {
                // Line delimiter parameter
                var delimiterPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };
                delimiterPanel.Children.Add(new TextBlock { Text = "Line Delimiter:", Width = 120, VerticalAlignment = VerticalAlignment.Center });
                delimiterPanel.Children.Add(new TextBox { Name = "paramLineDelimiter", Width = 300, Text = "" });
                pnlParameters.Children.Add(delimiterPanel);

                // Explanation text
                var explanationPanel = new StackPanel { Margin = new Thickness(0, 5, 0, 5) };
                explanationPanel.Children.Add(new TextBlock
                {
                    Text = "Leave the delimiter empty to treat each line as a single column. Enter a delimiter (like a comma) to split each line into multiple columns.",
                    TextWrapping = TextWrapping.Wrap
                });
                pnlParameters.Children.Add(explanationPanel);

                // First line as header parameter
                var headerPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 5) };
                headerPanel.Children.Add(new TextBlock { Text = "First Line is Header:", Width = 120, VerticalAlignment = VerticalAlignment.Center });
                headerPanel.Children.Add(new CheckBox { Name = "paramFirstLineHeader", IsChecked = false });
                pnlParameters.Children.Add(headerPanel);
            }
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

        private ConversionParameters GetConversionParameters()
        {
            var parameters = new ConversionParameters();

            // Extract parameters from the UI based on the selected conversion
            if (IsTextToHtmlSelected())
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
            else if (IsTxtToTsvSelected())
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

            return parameters;
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

                // Prepare for conversion
                btnConvert.IsEnabled = false;
                txtLog.Clear();
                progressBar.Value = 0;

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

                // Handle result
                if (result.Success)
                {
                    LogMessage($"Conversion successful! Time: {result.ElapsedTime.TotalSeconds:F2} seconds");
                    MessageBox.Show(
                        $"Conversion completed successfully in {result.ElapsedTime.TotalSeconds:F2} seconds.\nOutput file: {Path.GetFullPath(txtOutputPath.Text)}",
                        "Conversion Complete",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    LogMessage($"Conversion failed: {result.Error?.Message}");
                    MessageBox.Show(
                        $"Conversion failed: {result.Error?.Message}",
                        "Conversion Failed",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error: {ex.Message}");
                MessageBox.Show(
                    $"An error occurred: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                btnConvert.IsEnabled = true;
            }
        }

        private void LogMessage(string message)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            txtLog.ScrollToEnd();
        }

        private void btnBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Input File",
                Filter = "All Files (*.*)|*.*",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                txtInputPath.Text = dialog.FileName;

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

            if (dialog.ShowDialog() == true)
            {
                // Extract just the directory path
                string? directoryPath = Path.GetDirectoryName(dialog.FileName);
                if (!string.IsNullOrEmpty(directoryPath))
                {
                    txtOutputPath.Tag = directoryPath;
                    UpdateOutputPath();
                }
            }
        }

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
    }
}
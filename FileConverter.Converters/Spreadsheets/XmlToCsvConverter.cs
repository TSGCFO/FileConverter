using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace FileConverter.Converters.Spreadsheets
{
    /// <summary>
    /// Converter implementation for converting XML files to CSV format.
    /// </summary>
    public class XmlToCsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Xml };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Csv };

        /// <summary>
        /// Converts an XML file to CSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input XML file.</param>
        /// <param name="outputPath">Path where the output CSV file will be saved.</param>
        /// <param name="parameters">Optional parameters for customizing the conversion.</param>
        /// <param name="progress">Interface for reporting progress.</param>
        /// <param name="cancellationToken">Token for monitoring cancellation requests.</param>
        /// <returns>A conversion result with details about the operation.</returns>
        public async Task<ConversionResult> ConvertAsync(
            string inputPath,
            string outputPath,
            ConversionParameters parameters,
            IProgress<ConversionProgress>? progress,
            CancellationToken cancellationToken)
        {
            var startTime = DateTime.Now;

            try
            {
                // Report initial progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 0,
                    StatusMessage = "Starting XML to CSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                string rootElementPath = parameters.GetParameter("rootElementPath", "");
                string csvDelimiter = parameters.GetParameter("csvDelimiter", ",");
                bool includeHeaders = parameters.GetParameter("includeHeaders", true);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading XML file..."
                });

                // Load the XML document
                XDocument doc = await Task.Run(() => XDocument.Load(inputPath), cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Find the collection of elements to process
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Analyzing XML structure..."
                });

                // Find the elements to process based on the specified path
                var elements = FindElementsToProcess(doc, rootElementPath);

                if (elements.Count == 0)
                {
                    throw new InvalidOperationException($"No elements found at path '{rootElementPath}'. Please specify a valid element path.");
                }

                // Extract all unique attribute and element names to use as columns
                var columnNames = ExtractColumnNames(elements);

                if (columnNames.Count == 0)
                {
                    throw new InvalidOperationException("No attributes or child elements found to convert to columns.");
                }

                // Generate CSV content
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 60,
                    StatusMessage = "Converting to CSV format..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                var csv = GenerateCsv(elements, columnNames, csvDelimiter, includeHeaders);

                // Write to output file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing CSV file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                await File.WriteAllTextAsync(outputPath, csv, Encoding.UTF8, cancellationToken);

                // Report completion
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 100,
                    StatusMessage = "Conversion complete!"
                });

                return new ConversionResult
                {
                    Success = true,
                    InputPath = inputPath,
                    OutputPath = outputPath,
                    InputFormat = FileFormat.Xml,
                    OutputFormat = FileFormat.Csv,
                    ElapsedTime = DateTime.Now - startTime
                };
            }
            catch (OperationCanceledException)
            {
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 0,
                    StatusMessage = "Conversion canceled."
                });

                throw; // Rethrow to let the engine handle it
            }
            catch (Exception ex)
            {
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 0,
                    StatusMessage = $"Error: {ex.Message}"
                });

                return new ConversionResult
                {
                    Success = false,
                    InputPath = inputPath,
                    OutputPath = outputPath,
                    InputFormat = FileFormat.Xml,
                    OutputFormat = FileFormat.Csv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Finds XML elements to process based on the specified path.
        /// </summary>
        /// <param name="doc">The XML document.</param>
        /// <param name="rootElementPath">Path to the root element(s) to process.</param>
        /// <returns>A list of XML elements to process.</returns>
        private List<XElement> FindElementsToProcess(XDocument doc, string rootElementPath)
        {
            // If no path specified, try to find a collection of elements at the document root
            if (string.IsNullOrWhiteSpace(rootElementPath))
            {
                // Get the most common element name at the first level
                var rootElement = doc.Root;
                if (rootElement == null)
                    return new List<XElement>();

                var firstLevelElements = rootElement.Elements().ToList();

                if (firstLevelElements.Count == 0)
                    return new List<XElement> { rootElement }; // Use the root itself if no children

                // Find the most common element name
                var mostCommonElementName = firstLevelElements
                    .GroupBy(e => e.Name)
                    .OrderByDescending(g => g.Count())
                    .First()
                    .Key;

                return firstLevelElements.Where(e => e.Name == mostCommonElementName).ToList();
            }
            else
            {
                // Parse the path and navigate to the elements
                string[] pathParts = rootElementPath.Split('/');
                IEnumerable<XElement> currentElements = doc.Root != null
                    ? new List<XElement> { doc.Root }
                    : new List<XElement>();

                foreach (var part in pathParts.Where(p => !string.IsNullOrWhiteSpace(p)))
                {
                    if (currentElements == null || !currentElements.Any())
                        break;

                    currentElements = currentElements.SelectMany(e => e.Elements(part));
                }

                return currentElements?.ToList() ?? new List<XElement>();
            }
        }

        /// <summary>
        /// Extracts all unique column names from the XML elements.
        /// </summary>
        /// <param name="elements">The XML elements to analyze.</param>
        /// <returns>A list of column names.</returns>
        private List<string> ExtractColumnNames(List<XElement> elements)
        {
            var columnNames = new HashSet<string>();

            // Add all attributes and child elements across all elements
            foreach (var element in elements)
            {
                // Add attributes
                foreach (var attribute in element.Attributes())
                {
                    columnNames.Add(attribute.Name.LocalName);
                }

                // Add child elements with simple content
                foreach (var child in element.Elements())
                {
                    if (!child.HasElements)
                    {
                        columnNames.Add(child.Name.LocalName);
                    }
                }
            }

            return columnNames.ToList();
        }

        /// <summary>
        /// Generates CSV content from XML elements.
        /// </summary>
        /// <param name="elements">The XML elements to convert.</param>
        /// <param name="columnNames">The column names to include.</param>
        /// <param name="delimiter">The CSV delimiter character.</param>
        /// <param name="includeHeaders">Whether to include headers in the CSV.</param>
        /// <returns>The generated CSV content.</returns>
        private string GenerateCsv(List<XElement> elements, List<string> columnNames, string delimiter, bool includeHeaders)
        {
            var sb = new StringBuilder();

            // Add headers
            if (includeHeaders)
            {
                sb.AppendLine(string.Join(delimiter, columnNames.Select(name => EscapeForCsv(name, delimiter))));
            }

            // Add data rows
            foreach (var element in elements)
            {
                var row = new List<string>();

                foreach (var column in columnNames)
                {
                    // Try to get attribute value
                    var attribute = element.Attribute(column);
                    if (attribute != null)
                    {
                        row.Add(attribute.Value);
                        continue;
                    }

                    // Try to get child element value
                    var childElement = element.Element(column);
                    if (childElement != null && !childElement.HasElements)
                    {
                        row.Add(childElement.Value);
                        continue;
                    }

                    // If not found, add empty value
                    row.Add(string.Empty);
                }

                sb.AppendLine(string.Join(delimiter, row.Select(value => EscapeForCsv(value, delimiter))));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Escapes special characters for CSV format.
        /// </summary>
        /// <param name="field">The field to escape.</param>
        /// <param name="delimiter">The CSV delimiter character.</param>
        /// <returns>The escaped field.</returns>
        private string EscapeForCsv(string field, string delimiter)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            bool needsQuoting = field.Contains(delimiter) || field.Contains("\"") || field.Contains("\r") || field.Contains("\n");

            if (!needsQuoting)
                return field;

            // Escape quotes by doubling them
            field = field.Replace("\"", "\"\"");

            // Surround with quotes
            return $"\"{field}\"";
        }
    }
}
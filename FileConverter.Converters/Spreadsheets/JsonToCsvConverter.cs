using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Converters.Spreadsheets
{
    /// <summary>
    /// Converter implementation for converting JSON files to CSV format.
    /// </summary>
    public class JsonToCsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Json };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Csv };

        /// <summary>
        /// Converts a JSON file to CSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input JSON file.</param>
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
                    StatusMessage = "Starting JSON to CSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                char csvDelimiter = parameters.GetParameter("csvDelimiter", ',');
                char csvQuote = parameters.GetParameter("csvQuote", '"');
                string rootElement = parameters.GetParameter("rootElement", string.Empty);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading JSON file..."
                });

                // Read JSON file
                string jsonContent = await File.ReadAllTextAsync(inputPath, cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Parse JSON
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Parsing JSON data..."
                });

                // Parse the JSON content using System.Text.Json
                using (JsonDocument document = JsonDocument.Parse(jsonContent))
                {
                    // Get the root element
                    JsonElement root = document.RootElement;

                    // If a specific root element is specified and it exists, use that
                    if (!string.IsNullOrEmpty(rootElement) && root.ValueKind == JsonValueKind.Object)
                    {
                        if (root.TryGetProperty(rootElement, out JsonElement property))
                        {
                            root = property;
                        }
                        else
                        {
                            throw new InvalidOperationException($"Root element '{rootElement}' not found in JSON.");
                        }
                    }

                    // Check if we have an array to work with
                    if (root.ValueKind != JsonValueKind.Array)
                    {
                        // If not an array, wrap the single object in a list
                        if (root.ValueKind == JsonValueKind.Object)
                        {
                            // Create a single-item list with the object
                            root = JsonDocument.Parse($"[{jsonContent}]").RootElement;
                        }
                        else
                        {
                            throw new InvalidOperationException("JSON must contain an array of objects or a single object.");
                        }
                    }

                    // Extract fields from the first object to determine CSV headers
                    var headers = new List<string>();
                    var allRows = new List<Dictionary<string, string>>();

                    // Process all objects
                    for (int i = 0; i < root.GetArrayLength(); i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        JsonElement item = root[i];

                        if (item.ValueKind != JsonValueKind.Object)
                        {
                            continue; // Skip non-object items
                        }

                        var row = new Dictionary<string, string>();

                        // Extract properties from the object
                        foreach (JsonProperty property in item.EnumerateObject())
                        {
                            string key = property.Name;
                            string value = FormatJsonValue(property.Value);

                            // Add to the list of headers if not already present
                            if (!headers.Contains(key))
                            {
                                headers.Add(key);
                            }

                            row[key] = value;
                        }

                        allRows.Add(row);

                        // Report progress
                        if (i % Math.Max(1, root.GetArrayLength() / 10) == 0)
                        {
                            int percentComplete = 40 + (i * 30 / root.GetArrayLength());
                            progress?.Report(new ConversionProgress
                            {
                                PercentComplete = percentComplete,
                                StatusMessage = $"Processing JSON item {i + 1} of {root.GetArrayLength()}..."
                            });
                        }
                    }

                    // Write CSV
                    progress?.Report(new ConversionProgress
                    {
                        PercentComplete = 70,
                        StatusMessage = "Writing CSV file..."
                    });

                    using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
                    {
                        // Write headers
                        string headerLine = string.Join(csvDelimiter, headers.Select(h =>
                            EscapeForCsv(h, csvDelimiter, csvQuote)));
                        await writer.WriteLineAsync(headerLine);

                        // Write rows
                        for (int i = 0; i < allRows.Count; i++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var row = allRows[i];
                            var rowValues = headers.Select(header =>
                                row.TryGetValue(header, out string? value) ? value ?? string.Empty : string.Empty);

                            string rowLine = string.Join(csvDelimiter, rowValues.Select(v =>
                                EscapeForCsv(v, csvDelimiter, csvQuote)));

                            await writer.WriteLineAsync(rowLine);

                            // Report progress
                            if (i % Math.Max(1, allRows.Count / 10) == 0 || i == allRows.Count - 1)
                            {
                                int percentComplete = 70 + (i * 30 / allRows.Count);
                                progress?.Report(new ConversionProgress
                                {
                                    PercentComplete = percentComplete,
                                    StatusMessage = $"Writing row {i + 1} of {allRows.Count}..."
                                });
                            }
                        }
                    }
                }

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
                    InputFormat = FileFormat.Json,
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
                    InputFormat = FileFormat.Json,
                    OutputFormat = FileFormat.Csv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Formats a JsonElement value as a string.
        /// </summary>
        /// <param name="element">The JsonElement to format.</param>
        /// <returns>The formatted string value.</returns>
        private string FormatJsonValue(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.String:
                    return element.GetString() ?? string.Empty;
                case JsonValueKind.Number:
                    return element.ToString();
                case JsonValueKind.True:
                    return "true";
                case JsonValueKind.False:
                    return "false";
                case JsonValueKind.Null:
                    return string.Empty;
                case JsonValueKind.Object:
                case JsonValueKind.Array:
                    return element.GetRawText(); // Use the raw JSON for objects and arrays
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Escapes special characters for CSV format.
        /// </summary>
        /// <param name="field">The field to escape.</param>
        /// <param name="csvDelimiter">The CSV delimiter character.</param>
        /// <param name="csvQuote">The CSV quote character.</param>
        /// <returns>The escaped field.</returns>
        private string EscapeForCsv(string field, char csvDelimiter, char csvQuote)
        {
            // Check if the field needs quoting
            bool needsQuoting = field.Contains(csvDelimiter) ||
                               field.Contains(csvQuote) ||
                               field.Contains('\n') ||
                               field.Contains('\r');

            if (!needsQuoting)
            {
                return field;
            }

            // Escape quotes by doubling them
            string escapedField = field.Replace(csvQuote.ToString(), csvQuote.ToString() + csvQuote.ToString());

            // Surround with quotes
            return csvQuote + escapedField + csvQuote;
        }
    }
}
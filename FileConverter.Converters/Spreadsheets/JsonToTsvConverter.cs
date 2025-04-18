using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Converters.Spreadsheets
{
    /// <summary>
    /// Converter implementation for converting JSON files to TSV format.
    /// </summary>
    public class JsonToTsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Json };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Tsv };

        /// <summary>
        /// Converts a JSON file to TSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input JSON file.</param>
        /// <param name="outputPath">Path where the output TSV file will be saved.</param>
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
                    StatusMessage = "Starting JSON to TSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                string arrayPath = parameters.GetParameter("arrayPath", "");
                bool includeHeaders = parameters.GetParameter("includeHeaders", true);
                int maxDepth = parameters.GetParameter("maxDepth", 5);
                string flattenSeparator = parameters.GetParameter("flattenSeparator", ".");

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading JSON file..."
                });

                // Read the JSON file
                string jsonContent = await File.ReadAllTextAsync(inputPath, cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Parse the JSON
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Parsing JSON structure..."
                });

                // Parse JSON and find the array to convert
                var jsonArray = await Task.Run(() => ParseJsonAndFindArray(jsonContent, arrayPath), cancellationToken);

                if (jsonArray == null || jsonArray.Count == 0)
                {
                    throw new InvalidOperationException("No suitable array found in the JSON file for conversion.");
                }

                // Extract column names from all objects
                var columns = ExtractColumns(jsonArray, maxDepth, flattenSeparator);

                if (columns.Count == 0)
                {
                    throw new InvalidOperationException("No properties found to convert to columns.");
                }

                // Generate TSV content
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 60,
                    StatusMessage = "Converting to TSV format..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                var tsv = GenerateTsv(jsonArray, columns, includeHeaders, maxDepth, flattenSeparator);

                // Write to output file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing TSV file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                await File.WriteAllTextAsync(outputPath, tsv, Encoding.UTF8, cancellationToken);

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
                    OutputFormat = FileFormat.Tsv,
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
                    OutputFormat = FileFormat.Tsv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Parses the JSON content and finds the array to convert.
        /// </summary>
        /// <param name="jsonContent">The JSON content to parse.</param>
        /// <param name="arrayPath">Path to the array in the JSON. Empty for auto-detection.</param>
        /// <returns>A list of JSON objects to convert.</returns>
        private List<JsonObject> ParseJsonAndFindArray(string jsonContent, string arrayPath)
        {
            using (JsonDocument doc = JsonDocument.Parse(jsonContent))
            {
                JsonElement root = doc.RootElement;

                // If path is specified, try to follow it
                if (!string.IsNullOrWhiteSpace(arrayPath))
                {
                    // Split the path and navigate through it
                    string[] pathSegments = arrayPath.Split('.');
                    JsonElement current = root;

                    foreach (var segment in pathSegments)
                    {
                        if (current.ValueKind == JsonValueKind.Object)
                        {
                            if (current.TryGetProperty(segment, out var property))
                            {
                                current = property;
                            }
                            else
                            {
                                throw new InvalidOperationException($"Path segment '{segment}' not found in JSON.");
                            }
                        }
                        else
                        {
                            throw new InvalidOperationException($"Cannot navigate to '{segment}' because parent is not an object.");
                        }
                    }

                    // Check if we found an array
                    if (current.ValueKind == JsonValueKind.Array)
                    {
                        return ConvertJsonArrayToObjects(current);
                    }
                    else if (current.ValueKind == JsonValueKind.Object)
                    {
                        // Single object - wrap in a list
                        var jsonObj = JsonSerializer.Deserialize<JsonObject>(current.GetRawText());
                        return jsonObj != null ? new List<JsonObject> { jsonObj } : new List<JsonObject>();
                    }
                    else
                    {
                        throw new InvalidOperationException("The specified path does not point to an array or object.");
                    }
                }
                else
                {
                    // Auto-detect: if root is an array, use it
                    if (root.ValueKind == JsonValueKind.Array)
                    {
                        return ConvertJsonArrayToObjects(root);
                    }

                    // If root is an object, look for the first array property
                    if (root.ValueKind == JsonValueKind.Object)
                    {
                        // First, check if the root object itself should be converted
                        bool hasSimpleProperties = false;
                        foreach (var property in root.EnumerateObject())
                        {
                            if (property.Value.ValueKind != JsonValueKind.Object &&
                                property.Value.ValueKind != JsonValueKind.Array)
                            {
                                hasSimpleProperties = true;
                                break;
                            }
                        }

                        if (hasSimpleProperties)
                        {
                            // Use the root object itself
                            var rootObj = JsonSerializer.Deserialize<JsonObject>(root.GetRawText());
                            return rootObj != null ? new List<JsonObject> { rootObj } : new List<JsonObject>();
                        }

                        // Look for arrays in the root object's properties
                        foreach (var property in root.EnumerateObject())
                        {
                            if (property.Value.ValueKind == JsonValueKind.Array && property.Value.GetArrayLength() > 0)
                            {
                                return ConvertJsonArrayToObjects(property.Value);
                            }
                        }

                        // If no arrays found, just use the root object
                        var fallbackObj = JsonSerializer.Deserialize<JsonObject>(root.GetRawText());
                        return fallbackObj != null ? new List<JsonObject> { fallbackObj } : new List<JsonObject>();
                    }

                    // Couldn't find a suitable array
                    throw new InvalidOperationException("No array found in the JSON file for conversion.");
                }
            }
        }

        /// <summary>
        /// Converts a JsonElement array to a list of JsonObjects.
        /// </summary>
        /// <param name="arrayElement">The JSON array element.</param>
        /// <returns>A list of JsonObjects.</returns>
        private List<JsonObject> ConvertJsonArrayToObjects(JsonElement arrayElement)
        {
            var result = new List<JsonObject>();

            foreach (var item in arrayElement.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.Object)
                {
                    var obj = JsonSerializer.Deserialize<JsonObject>(item.GetRawText());
                    if (obj != null)
                    {
                        result.Add(obj);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Extracts all columns from the JSON objects.
        /// </summary>
        /// <param name="objects">The JSON objects to analyze.</param>
        /// <param name="maxDepth">Maximum depth for nested objects.</param>
        /// <param name="separator">Separator for flattened property names.</param>
        /// <returns>A list of column names.</returns>
        private List<string> ExtractColumns(List<JsonObject> objects, int maxDepth, string separator)
        {
            var columns = new HashSet<string>();

            foreach (var obj in objects)
            {
                // Extract properties from this object with flattening
                ExtractPropertiesRecursive(obj, "", columns, maxDepth, separator, 0);
            }

            return columns.ToList();
        }

        /// <summary>
        /// Recursively extracts properties from a JSON object and adds them to the columns set.
        /// </summary>
        /// <param name="obj">The JSON object to process.</param>
        /// <param name="prefix">Prefix for nested properties.</param>
        /// <param name="columns">Set of column names to add to.</param>
        /// <param name="maxDepth">Maximum recursion depth.</param>
        /// <param name="separator">Separator for flattened property names.</param>
        /// <param name="currentDepth">Current recursion depth.</param>
        private void ExtractPropertiesRecursive(JsonObject obj, string prefix, HashSet<string> columns,
            int maxDepth, string separator, int currentDepth)
        {
            if (currentDepth >= maxDepth || obj == null)
                return;

            foreach (var property in obj)
            {
                string propName = string.IsNullOrEmpty(prefix) ? property.Key : prefix + separator + property.Key;

                if (property.Value is JsonObject nestedObj)
                {
                    // Recurse into nested object
                    ExtractPropertiesRecursive(nestedObj, propName, columns, maxDepth, separator, currentDepth + 1);
                }
                else if (property.Value is JsonArray array)
                {
                    // For arrays, we'll just use the property name and stringify the array
                    columns.Add(propName);
                }
                else
                {
                    // Simple value
                    columns.Add(propName);
                }
            }
        }

        /// <summary>
        /// Generates TSV content from JSON objects.
        /// </summary>
        /// <param name="objects">The JSON objects to convert.</param>
        /// <param name="columns">The column names to include.</param>
        /// <param name="includeHeaders">Whether to include headers in the TSV.</param>
        /// <param name="maxDepth">Maximum depth for nested objects.</param>
        /// <param name="separator">Separator for flattened property names.</param>
        /// <returns>The generated TSV content.</returns>
        private string GenerateTsv(List<JsonObject> objects, List<string> columns, bool includeHeaders,
            int maxDepth, string separator)
        {
            var sb = new StringBuilder();

            // Add headers
            if (includeHeaders)
            {
                sb.AppendLine(string.Join("\t", columns.Select(name => EscapeForTsv(name))));
            }

            // Add data rows
            foreach (var obj in objects)
            {
                var values = new List<string>();

                foreach (var column in columns)
                {
                    string value = ExtractValue(obj, column, separator);
                    values.Add(EscapeForTsv(value));
                }

                sb.AppendLine(string.Join("\t", values));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Extracts a value from a JSON object using the specified property path.
        /// </summary>
        /// <param name="obj">The JSON object to extract from.</param>
        /// <param name="propertyPath">The property path (dot-separated for nested properties).</param>
        /// <param name="separator">Separator for flattened property names.</param>
        /// <returns>The extracted value as a string.</returns>
        private string ExtractValue(JsonObject obj, string propertyPath, string separator)
        {
            string[] pathParts = propertyPath.Split(new[] { separator }, StringSplitOptions.None);
            JsonNode? currentNode = obj;

            for (int i = 0; i < pathParts.Length; i++)
            {
                var part = pathParts[i];

                if (currentNode is JsonObject currentObj)
                {
                    if (currentObj.TryGetPropertyValue(part, out var value) && value != null)
                    {
                        currentNode = value;
                    }
                    else
                    {
                        return string.Empty; // Property not found or value is null
                    }
                }
                else
                {
                    return string.Empty; // Not an object
                }
            }

            // Convert the final node to a string
            if (currentNode == null)
            {
                return string.Empty;
            }
            else if (currentNode is JsonObject)
            {
                return JsonSerializer.Serialize(currentNode);
            }
            else if (currentNode is JsonArray)
            {
                return JsonSerializer.Serialize(currentNode);
            }
            else
            {
                return currentNode.GetValue<JsonElement>().ToString();
            }
        }

        /// <summary>
        /// Escapes special characters for TSV format.
        /// </summary>
        /// <param name="field">The field to escape.</param>
        /// <returns>The escaped field.</returns>
        private string EscapeForTsv(string field)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            // Replace tabs with spaces to avoid breaking TSV structure
            return field.Replace("\t", " ");
        }
    }
}
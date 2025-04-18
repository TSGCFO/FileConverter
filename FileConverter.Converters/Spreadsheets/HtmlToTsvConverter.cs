using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Converters.Spreadsheets
{
    /// <summary>
    /// Converter implementation for extracting tabular data from HTML files and converting to TSV format.
    /// </summary>
    public class HtmlToTsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Html };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Tsv };

        /// <summary>
        /// Converts an HTML file to TSV format by extracting tabular data.
        /// </summary>
        /// <param name="inputPath">Path to the input HTML file.</param>
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
                    StatusMessage = "Starting HTML to TSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                int tableIndex = parameters.GetParameter("tableIndex", 0); // Which table to extract (0 = first)
                bool includeHeaders = parameters.GetParameter("includeHeaders", true);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading HTML file..."
                });

                // Read HTML file
                string htmlContent = await File.ReadAllTextAsync(inputPath, cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Extract tables from HTML
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Extracting tables from HTML..."
                });

                var tables = ExtractTablesFromHtml(htmlContent);

                if (tables.Count == 0)
                {
                    throw new InvalidOperationException("No tables found in the HTML file.");
                }

                if (tableIndex >= tables.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(tableIndex), $"Table index {tableIndex} is out of range. Only {tables.Count} tables found.");
                }

                // Get the selected table
                var selectedTable = tables[tableIndex];

                // Convert table to TSV
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 60,
                    StatusMessage = "Converting table to TSV format..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                StringBuilder tsvBuilder = new StringBuilder();
                int startRow = includeHeaders ? 0 : 1;

                for (int i = startRow; i < selectedTable.Count; i++)
                {
                    tsvBuilder.AppendLine(string.Join("\t", selectedTable[i].Select(cell => EscapeForTsv(cell))));
                }

                // Write the TSV file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing TSV file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                await File.WriteAllTextAsync(outputPath, tsvBuilder.ToString(), Encoding.UTF8, cancellationToken);

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
                    InputFormat = FileFormat.Html,
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
                    InputFormat = FileFormat.Html,
                    OutputFormat = FileFormat.Tsv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts tables from HTML content.
        /// </summary>
        /// <param name="htmlContent">The HTML content to process.</param>
        /// <returns>A list of tables, where each table is a list of rows, and each row is a list of cells.</returns>
        private List<List<List<string>>> ExtractTablesFromHtml(string htmlContent)
        {
            var tables = new List<List<List<string>>>();

            // Simple regex-based approach to extract tables
            var tablePattern = @"<table[^>]*>(.*?)</table>";
            var rowPattern = @"<tr[^>]*>(.*?)</tr>";
            var cellPattern = @"<t[hd][^>]*>(.*?)</t[hd]>";

            var tableMatches = Regex.Matches(htmlContent, tablePattern, RegexOptions.Singleline);

            foreach (Match tableMatch in tableMatches)
            {
                var table = new List<List<string>>();
                var rowMatches = Regex.Matches(tableMatch.Groups[1].Value, rowPattern, RegexOptions.Singleline);

                foreach (Match rowMatch in rowMatches)
                {
                    var row = new List<string>();
                    var cellMatches = Regex.Matches(rowMatch.Groups[1].Value, cellPattern, RegexOptions.Singleline);

                    foreach (Match cellMatch in cellMatches)
                    {
                        // Clean up the cell content (remove HTML tags, decode entities, etc.)
                        string cellContent = CleanHtmlContent(cellMatch.Groups[1].Value);
                        row.Add(cellContent);
                    }

                    if (row.Count > 0)
                    {
                        table.Add(row);
                    }
                }

                if (table.Count > 0)
                {
                    tables.Add(table);
                }
            }

            return tables;
        }

        /// <summary>
        /// Cleans HTML content by removing tags and decoding entities.
        /// </summary>
        /// <param name="html">The HTML content to clean.</param>
        /// <returns>The cleaned content.</returns>
        private string CleanHtmlContent(string html)
        {
            // Remove HTML tags
            var withoutTags = Regex.Replace(html, @"<[^>]+>", string.Empty);

            // Decode HTML entities like &nbsp;, &quot;, etc.
            withoutTags = System.Net.WebUtility.HtmlDecode(withoutTags);

            // Normalize whitespace
            withoutTags = Regex.Replace(withoutTags, @"\s+", " ").Trim();

            return withoutTags;
        }

        /// <summary>
        /// Escapes special characters for TSV format.
        /// </summary>
        /// <param name="field">The field to escape.</param>
        /// <returns>The escaped field.</returns>
        private string EscapeForTsv(string field)
        {
            // Replace tabs with spaces to avoid breaking TSV structure
            return field.Replace("\t", " ");
        }
    }
}
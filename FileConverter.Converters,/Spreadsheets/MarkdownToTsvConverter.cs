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
    /// Converter implementation for extracting tables from Markdown files and converting to TSV format.
    /// </summary>
    public class MarkdownToTsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Md };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Tsv };

        /// <summary>
        /// Converts a Markdown file to TSV format by extracting tabular data.
        /// </summary>
        /// <param name="inputPath">Path to the input Markdown file.</param>
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
                    StatusMessage = "Starting Markdown to TSV conversion..."
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
                    StatusMessage = "Reading Markdown file..."
                });

                // Read Markdown file
                string markdownContent = await File.ReadAllTextAsync(inputPath, cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Extract tables from Markdown
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Extracting tables from Markdown..."
                });

                var tables = ExtractTablesFromMarkdown(markdownContent);

                if (tables.Count == 0)
                {
                    throw new InvalidOperationException("No tables found in the Markdown file.");
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
                    // Skip the separator row (row 1) in Markdown tables
                    if (i == 1 && includeHeaders)
                        continue;

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
                    InputFormat = FileFormat.Md,
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
                    InputFormat = FileFormat.Md,
                    OutputFormat = FileFormat.Tsv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts tables from Markdown content.
        /// </summary>
        /// <param name="markdownContent">The Markdown content to process.</param>
        /// <returns>A list of tables, where each table is a list of rows, and each row is a list of cells.</returns>
        private List<List<List<string>>> ExtractTablesFromMarkdown(string markdownContent)
        {
            var tables = new List<List<List<string>>>();

            // Split content by lines
            string[] lines = markdownContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            List<List<string>>? currentTable = null;
            bool isInTable = false;

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();

                // Check if this is a table row
                if (trimmedLine.StartsWith("|") && trimmedLine.EndsWith("|"))
                {
                    // If not already in a table, start a new one
                    if (!isInTable)
                    {
                        currentTable = new List<List<string>>();
                        isInTable = true;
                    }

                    // Extract cells from the row
                    string rowContent = trimmedLine.Substring(1, trimmedLine.Length - 2);
                    string[] cells = rowContent.Split('|');

                    // Add the row to the current table
                    if (currentTable != null)
                    {
                        currentTable.Add(cells.Select(cell => cell.Trim()).ToList());
                    }
                }
                else if (isInTable)
                {
                    // We've reached the end of a table
                    isInTable = false;

                    // Add the completed table to the list
                    if (currentTable != null && currentTable.Count > 0)
                    {
                        tables.Add(currentTable);
                        currentTable = null;
                    }
                }
            }

            // Don't forget to add the last table if we reached the end of the file
            if (isInTable && currentTable != null && currentTable.Count > 0)
            {
                tables.Add(currentTable);
            }

            return tables;
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
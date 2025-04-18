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

namespace FileConverter.Converters.Spreadsheets
{
    /// <summary>
    /// Converter implementation for converting CSV files to TSV format.
    /// </summary>
    public class CsvToTsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Csv };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Tsv };

        /// <summary>
        /// Converts a CSV file to TSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input CSV file.</param>
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
                    StatusMessage = "Starting CSV to TSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get CSV parameters
                char csvDelimiter = parameters.GetParameter("csvDelimiter", ',');
                char csvQuote = parameters.GetParameter("csvQuote", '"');
                bool hasHeader = parameters.GetParameter("hasHeader", true);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 10,
                    StatusMessage = "Reading CSV file..."
                });

                // Read all lines from the CSV file
                string[] lines = await File.ReadAllLinesAsync(inputPath, cancellationToken);
                int totalLines = lines.Length;

                if (totalLines == 0)
                {
                    // Write empty file and return success
                    await File.WriteAllTextAsync(outputPath, string.Empty, cancellationToken);

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
                        InputFormat = FileFormat.Csv,
                        OutputFormat = FileFormat.Tsv,
                        ElapsedTime = DateTime.Now - startTime
                    };
                }

                // Prepare to write TSV
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Converting to TSV format..."
                });

                // Process in batches for better progress reporting
                using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
                {
                    for (int i = 0; i < totalLines; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string line = lines[i];
                        string tsvLine = ConvertCsvLineToTsv(line, csvDelimiter, csvQuote);
                        await writer.WriteLineAsync(tsvLine);

                        // Report progress periodically
                        if (i % Math.Max(1, totalLines / 10) == 0 || i == totalLines - 1)
                        {
                            int percentComplete = 40 + (i * 55 / totalLines);
                            progress?.Report(new ConversionProgress
                            {
                                PercentComplete = percentComplete,
                                StatusMessage = $"Converting line {i + 1} of {totalLines}..."
                            });
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
                    InputFormat = FileFormat.Csv,
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
                    InputFormat = FileFormat.Csv,
                    OutputFormat = FileFormat.Tsv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Converts a CSV line to TSV format.
        /// </summary>
        /// <param name="csvLine">The CSV line to convert.</param>
        /// <param name="csvDelimiter">The CSV delimiter character.</param>
        /// <param name="csvQuote">The CSV quote character.</param>
        /// <returns>The line converted to TSV format.</returns>
        private string ConvertCsvLineToTsv(string csvLine, char csvDelimiter, char csvQuote)
        {
            var fields = new List<string>();
            var field = new StringBuilder();
            bool inQuotes = false;

            // Parse CSV fields
            for (int i = 0; i < csvLine.Length; i++)
            {
                char c = csvLine[i];

                if (c == csvQuote)
                {
                    if (inQuotes && i + 1 < csvLine.Length && csvLine[i + 1] == csvQuote)
                    {
                        // Escaped quote
                        field.Append(csvQuote);
                        i++; // Skip the next quote
                    }
                    else
                    {
                        // Toggle quote state
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == csvDelimiter && !inQuotes)
                {
                    // End of field
                    fields.Add(field.ToString());
                    field.Clear();
                }
                else
                {
                    field.Append(c);
                }
            }

            // Add the last field
            fields.Add(field.ToString());

            // Convert to TSV
            return string.Join('\t', fields.Select(f => EscapeForTsv(f)));
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
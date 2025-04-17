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
    /// Converter implementation for converting plain text files to CSV format.
    /// </summary>
    public class TxtToCsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Txt };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Csv };

        /// <summary>
        /// Converts a plain text file to CSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input text file.</param>
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
                    StatusMessage = "Starting text to CSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                string lineDelimiter = parameters.GetParameter("lineDelimiter", string.Empty);
                bool treatFirstLineAsHeader = parameters.GetParameter("treatFirstLineAsHeader", false);
                char csvDelimiter = parameters.GetParameter("csvDelimiter", ',');
                char csvQuote = parameters.GetParameter("csvQuote", '"');

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading text file..."
                });

                // Read all lines from the text file
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
                        InputFormat = FileFormat.Txt,
                        OutputFormat = FileFormat.Csv,
                        ElapsedTime = DateTime.Now - startTime
                    };
                }

                // Prepare to write CSV
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 40,
                    StatusMessage = "Converting to CSV format..."
                });

                using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
                {
                    for (int i = 0; i < totalLines; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // Skip empty lines
                        if (string.IsNullOrWhiteSpace(lines[i]))
                            continue;

                        string csvLine;

                        // If a delimiter is specified, split the line into columns
                        if (!string.IsNullOrEmpty(lineDelimiter))
                        {
                            string[] fields = lines[i].Split(lineDelimiter);
                            var csvFields = fields.Select(field =>
                                EscapeForCsv(field.Trim(), csvDelimiter, csvQuote));
                            csvLine = string.Join(csvDelimiter, csvFields);
                        }
                        else
                        {
                            // Otherwise, treat the entire line as a single column
                            csvLine = EscapeForCsv(lines[i], csvDelimiter, csvQuote);
                        }

                        await writer.WriteLineAsync(csvLine);

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
                    InputFormat = FileFormat.Txt,
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
                    InputFormat = FileFormat.Txt,
                    OutputFormat = FileFormat.Csv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
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
using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Converters.Documents
{
    /// <summary>
    /// Converter implementation for converting DOCX files to CSV format.
    /// </summary>
    public class DocxToCsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Docx };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Csv };

        /// <summary>
        /// Converts a DOCX file to CSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input DOCX file.</param>
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
                    StatusMessage = "Starting DOCX to CSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                int tableIndex = parameters.GetParameter("tableIndex", 0); // Which table to extract (0 = first)
                bool includeHeaders = parameters.GetParameter("includeHeaders", true);
                char csvDelimiter = parameters.GetParameter("csvDelimiter", ',');
                char csvQuote = parameters.GetParameter("csvQuote", '"');

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading DOCX file..."
                });

                // Extract tables from the DOCX file
                var tables = await Task.Run(() => ExtractTablesFromDocx(inputPath), cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                if (tables.Count == 0)
                {
                    throw new InvalidOperationException("No tables found in the DOCX file.");
                }

                if (tableIndex >= tables.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(tableIndex),
                        $"Table index {tableIndex} is out of range. Only {tables.Count} tables found.");
                }

                // Get the selected table
                var selectedTable = tables[tableIndex];

                // Convert table to CSV
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 60,
                    StatusMessage = "Converting table to CSV format..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                var csvContent = ConvertTableToCsv(selectedTable, includeHeaders, csvDelimiter, csvQuote);

                // Write the CSV file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing CSV file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                await File.WriteAllTextAsync(outputPath, csvContent, Encoding.UTF8, cancellationToken);

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
                    InputFormat = FileFormat.Docx,
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
                    InputFormat = FileFormat.Docx,
                    OutputFormat = FileFormat.Csv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts tables from a DOCX file.
        /// </summary>
        /// <param name="docxPath">Path to the DOCX file.</param>
        /// <returns>A list of tables, where each table is a list of rows, and each row is a list of cells.</returns>
        private List<List<List<string>>> ExtractTablesFromDocx(string docxPath)
        {
            var result = new List<List<List<string>>>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
            {
                var body = doc.MainDocumentPart?.Document.Body;
                if (body == null)
                    return result;

                // Find all tables in the document
                var tables = body.Descendants<Table>();

                foreach (var table in tables)
                {
                    var tableData = new List<List<string>>();

                    // Process rows in the table
                    foreach (var row in table.Elements<TableRow>())
                    {
                        var rowData = new List<string>();

                        // Process cells in the row
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            // Extract text from the cell
                            string cellText = string.Join(" ", cell.Descendants<Text>().Select(t => t.Text));
                            rowData.Add(cellText);
                        }

                        if (rowData.Count > 0)
                        {
                            tableData.Add(rowData);
                        }
                    }

                    if (tableData.Count > 0)
                    {
                        result.Add(tableData);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Converts a table to CSV format.
        /// </summary>
        /// <param name="table">The table to convert.</param>
        /// <param name="includeHeaders">Whether to include headers in the CSV.</param>
        /// <param name="csvDelimiter">The CSV delimiter character.</param>
        /// <param name="csvQuote">The CSV quote character.</param>
        /// <returns>The CSV content.</returns>
        private string ConvertTableToCsv(
            List<List<string>> table,
            bool includeHeaders,
            char csvDelimiter,
            char csvQuote)
        {
            var sb = new StringBuilder();
            int startRow = includeHeaders ? 0 : 1;

            // Check if there are enough rows when headers are included
            if (includeHeaders && table.Count < 2)
            {
                // If there's only one row and headers are required, use it as data not header
                startRow = 0;
                includeHeaders = false;
            }

            // Process rows
            for (int i = startRow; i < table.Count; i++)
            {
                var row = table[i];

                // Convert each cell in the row to CSV format
                var csvRow = string.Join(
                    csvDelimiter,
                    row.Select(cell => EscapeForCsv(cell, csvDelimiter, csvQuote))
                );

                sb.AppendLine(csvRow);
            }

            return sb.ToString();
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
            if (string.IsNullOrEmpty(field))
                return string.Empty;

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
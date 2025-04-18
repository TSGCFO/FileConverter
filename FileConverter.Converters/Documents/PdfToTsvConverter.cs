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
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace FileConverter.Converters.Documents
{
    /// <summary>
    /// Converter implementation for extracting tables from PDF files and converting to TSV format.
    /// </summary>
    public class PdfToTsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Pdf };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Tsv };

        /// <summary>
        /// Converts a PDF file to TSV format by extracting tables.
        /// </summary>
        /// <param name="inputPath">Path to the input PDF file.</param>
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
                    StatusMessage = "Starting PDF to TSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                int pageNumber = parameters.GetParameter("pageNumber", 1); // Default to first page
                int tableIndex = parameters.GetParameter("tableIndex", 0); // Which table to extract (0 = first)
                bool includeHeaders = parameters.GetParameter("includeHeaders", true);
                double tableDetectionThreshold = parameters.GetParameter("tableDetectionThreshold", 10.0); // Gap threshold for table detection

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading PDF file..."
                });

                // Extract tables from the PDF file
                var tables = await Task.Run(() => ExtractTablesFromPdf(inputPath, pageNumber, tableDetectionThreshold), cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                if (tables.Count == 0)
                {
                    throw new InvalidOperationException("No tables found in the PDF file. Try adjusting the table detection threshold.");
                }

                if (tableIndex >= tables.Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(tableIndex),
                        $"Table index {tableIndex} is out of range. Only {tables.Count} tables found.");
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

                var tsvContent = ConvertTableToTsv(selectedTable, includeHeaders);

                // Write the TSV file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing TSV file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                await File.WriteAllTextAsync(outputPath, tsvContent, Encoding.UTF8, cancellationToken);

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
                    InputFormat = FileFormat.Pdf,
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
                    InputFormat = FileFormat.Pdf,
                    OutputFormat = FileFormat.Tsv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts tables from a PDF file.
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file.</param>
        /// <param name="pageNumber">The page number to extract tables from (1-based).</param>
        /// <param name="lineGapThreshold">Threshold for detecting table rows.</param>
        /// <returns>A list of tables, where each table is a list of rows, and each row is a list of cells.</returns>
        private List<List<List<string>>> ExtractTablesFromPdf(string pdfPath, int pageNumber, double lineGapThreshold)
        {
            var result = new List<List<List<string>>>();

            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                // Validate page number
                if (pageNumber < 1 || pageNumber > document.NumberOfPages)
                {
                    throw new ArgumentOutOfRangeException(nameof(pageNumber),
                        $"Page number {pageNumber} is out of range. PDF has {document.NumberOfPages} pages.");
                }

                // Get the requested page
                Page page = document.GetPage(pageNumber);

                // Get all words on the page
                var words = page.GetWords().ToList();

                if (words.Count == 0)
                    return result;

                // Group words by line (based on their Y position)
                var lines = GroupWordsByLines(words, lineGapThreshold);

                // Try to detect table structure
                var table = DetectTableStructure(lines, lineGapThreshold);

                if (table.Count > 0)
                {
                    result.Add(table);
                }
            }

            return result;
        }

        /// <summary>
        /// Groups words by lines based on their Y position.
        /// </summary>
        /// <param name="words">The words to group.</param>
        /// <param name="lineGapThreshold">The threshold for determining line breaks.</param>
        /// <returns>A list of lines, where each line is a list of words.</returns>
        private List<List<Word>> GroupWordsByLines(List<Word> words, double lineGapThreshold)
        {
            var lines = new List<List<Word>>();

            // Sort words by Y position (descending) and then by X position (ascending)
            var sortedWords = words.OrderByDescending(w => w.BoundingBox.Bottom)
                                  .ThenBy(w => w.BoundingBox.Left)
                                  .ToList();

            if (sortedWords.Count == 0)
                return lines;

            // Initialize the first line
            var currentLine = new List<Word> { sortedWords[0] };
            double currentY = sortedWords[0].BoundingBox.Bottom;

            // Process remaining words
            for (int i = 1; i < sortedWords.Count; i++)
            {
                var word = sortedWords[i];
                double yDiff = Math.Abs(word.BoundingBox.Bottom - currentY);

                // If the word is on the same line (within the threshold)
                if (yDiff <= lineGapThreshold)
                {
                    currentLine.Add(word);
                }
                else
                {
                    // Sort words in the line by X position
                    currentLine = currentLine.OrderBy(w => w.BoundingBox.Left).ToList();
                    lines.Add(currentLine);

                    // Start a new line
                    currentLine = new List<Word> { word };
                    currentY = word.BoundingBox.Bottom;
                }
            }

            // Add the last line
            if (currentLine.Count > 0)
            {
                currentLine = currentLine.OrderBy(w => w.BoundingBox.Left).ToList();
                lines.Add(currentLine);
            }

            return lines;
        }

        /// <summary>
        /// Detects table structure from lines of words.
        /// </summary>
        /// <param name="lines">The lines of words.</param>
        /// <param name="columnGapThreshold">The threshold for determining column breaks.</param>
        /// <returns>A table structure as a list of rows, where each row is a list of cell values.</returns>
        private List<List<string>> DetectTableStructure(List<List<Word>> lines, double columnGapThreshold)
        {
            var table = new List<List<string>>();

            if (lines.Count == 0)
                return table;

            // Detect column positions by analyzing word positions across all lines
            var columnBoundaries = DetectColumnBoundaries(lines, columnGapThreshold);

            // Process each line to form table rows
            foreach (var line in lines)
            {
                // Skip lines that are likely not part of the table (too few words)
                if (line.Count < 2)
                    continue;

                var row = new List<string>();

                // Assign words to columns based on their position
                var cellContents = new Dictionary<int, List<Word>>();

                foreach (var word in line)
                {
                    int columnIndex = DetermineColumnIndex(word, columnBoundaries);

                    if (!cellContents.ContainsKey(columnIndex))
                    {
                        cellContents[columnIndex] = new List<Word>();
                    }

                    cellContents[columnIndex].Add(word);
                }

                // Create cell values for each column
                for (int i = 0; i < columnBoundaries.Count - 1; i++)
                {
                    if (cellContents.TryGetValue(i, out var columnWords))
                    {
                        // Combine words in this column to form the cell value
                        string cellValue = string.Join(" ", columnWords.OrderBy(w => w.BoundingBox.Left)
                                                                    .Select(w => w.Text));
                        row.Add(cellValue);
                    }
                    else
                    {
                        // Empty cell
                        row.Add(string.Empty);
                    }
                }

                if (row.Count > 0)
                {
                    table.Add(row);
                }
            }

            // Filter out rows that don't match the most common column count
            if (table.Count > 0)
            {
                var columnCounts = table.GroupBy(r => r.Count)
                                       .OrderByDescending(g => g.Count())
                                       .ToList();

                if (columnCounts.Count > 0)
                {
                    int mostCommonColumnCount = columnCounts[0].Key;
                    table = table.Where(r => r.Count == mostCommonColumnCount).ToList();
                }
            }

            return table;
        }

        /// <summary>
        /// Detects column boundaries by analyzing word positions.
        /// </summary>
        /// <param name="lines">The lines of words.</param>
        /// <param name="columnGapThreshold">The threshold for determining column breaks.</param>
        /// <returns>A list of X positions representing column boundaries.</returns>
        private List<double> DetectColumnBoundaries(List<List<Word>> lines, double columnGapThreshold)
        {
            // Get all word boundary positions
            var allXPositions = new List<double>();

            foreach (var line in lines)
            {
                foreach (var word in line)
                {
                    allXPositions.Add(word.BoundingBox.Left);
                    allXPositions.Add(word.BoundingBox.Right);
                }
            }

            // Sort positions
            allXPositions.Sort();

            // Identify gaps that might indicate column boundaries
            var columnBoundaries = new List<double> { 0 }; // Start with left boundary

            for (int i = 1; i < allXPositions.Count; i++)
            {
                double gap = allXPositions[i] - allXPositions[i - 1];

                if (gap >= columnGapThreshold)
                {
                    double boundary = (allXPositions[i] + allXPositions[i - 1]) / 2;
                    columnBoundaries.Add(boundary);
                }
            }

            // Add right boundary if not already present
            if (allXPositions.Count > 0)
            {
                columnBoundaries.Add(allXPositions.Max() + 1);
            }

            // Ensure we have a reasonable number of columns
            if (columnBoundaries.Count < 3) // Need at least 2 columns (3 boundaries)
            {
                // Fallback to a simple approach
                double minX = allXPositions.Min();
                double maxX = allXPositions.Max();
                double width = maxX - minX;

                columnBoundaries.Clear();
                columnBoundaries.Add(minX);
                columnBoundaries.Add(minX + width / 3);
                columnBoundaries.Add(minX + 2 * width / 3);
                columnBoundaries.Add(maxX);
            }

            return columnBoundaries;
        }

        /// <summary>
        /// Determines the column index for a word based on column boundaries.
        /// </summary>
        /// <param name="word">The word to check.</param>
        /// <param name="columnBoundaries">The column boundaries.</param>
        /// <returns>The column index for the word.</returns>
        private int DetermineColumnIndex(Word word, List<double> columnBoundaries)
        {
            double wordCenter = (word.BoundingBox.Left + word.BoundingBox.Right) / 2;

            for (int i = 0; i < columnBoundaries.Count - 1; i++)
            {
                if (wordCenter >= columnBoundaries[i] && wordCenter < columnBoundaries[i + 1])
                {
                    return i;
                }
            }

            // Fallback to the last column
            return columnBoundaries.Count - 2;
        }

        /// <summary>
        /// Converts a table to TSV format.
        /// </summary>
        /// <param name="table">The table to convert.</param>
        /// <param name="includeHeaders">Whether to include headers in the TSV.</param>
        /// <returns>The TSV content.</returns>
        private string ConvertTableToTsv(List<List<string>> table, bool includeHeaders)
        {
            var sb = new StringBuilder();
            int startRow = includeHeaders ? 0 : 1;

            // Check if there are enough rows when headers are included
            if (includeHeaders && table.Count < 2)
            {
                startRow = 0;
                includeHeaders = false;
            }

            // Process rows
            for (int i = startRow; i < table.Count; i++)
            {
                var row = table[i];

                // Convert each cell in the row to TSV format
                var tsvRow = string.Join(
                    "\t",
                    row.Select(cell => EscapeForTsv(cell))
                );

                sb.AppendLine(tsvRow);
            }

            return sb.ToString();
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
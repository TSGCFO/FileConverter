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
    /// Converter implementation for extracting text content from PDF files.
    /// </summary>
    public class PdfToTxtConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Pdf };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Txt };

        /// <summary>
        /// Converts a PDF file to text format by extracting its content.
        /// </summary>
        /// <param name="inputPath">Path to the input PDF file.</param>
        /// <param name="outputPath">Path where the output text file will be saved.</param>
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
                    StatusMessage = "Starting PDF to text conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                bool preservePageBreaks = parameters.GetParameter("preservePageBreaks", true);
                bool includePageNumbers = parameters.GetParameter("includePageNumbers", false);
                bool orderByPosition = parameters.GetParameter("orderByPosition", true);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading PDF file..."
                });

                // Extract text from PDF
                var extractedText = await Task.Run(() =>
                    ExtractTextFromPdf(inputPath, preservePageBreaks, includePageNumbers, orderByPosition),
                    cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Write the text file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing text file..."
                });

                await File.WriteAllTextAsync(outputPath, extractedText, Encoding.UTF8, cancellationToken);

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
                    OutputFormat = FileFormat.Txt,
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
                    OutputFormat = FileFormat.Txt,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts text content from a PDF file.
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file.</param>
        /// <param name="preservePageBreaks">Whether to insert page break markers between pages.</param>
        /// <param name="includePageNumbers">Whether to include page numbers in the output.</param>
        /// <param name="orderByPosition">Whether to order text by position on the page.</param>
        /// <returns>The extracted text content.</returns>
        private string ExtractTextFromPdf(
            string pdfPath,
            bool preservePageBreaks,
            bool includePageNumbers,
            bool orderByPosition)
        {
            var sb = new StringBuilder();

            using (PdfDocument document = PdfDocument.Open(pdfPath))
            {
                for (int i = 0; i < document.NumberOfPages; i++)
                {
                    // Get the current page (1-based indexing for PdfPig)
                    Page page = document.GetPage(i + 1);

                    // Add page number if requested
                    if (includePageNumbers)
                    {
                        sb.AppendLine($"--- Page {i + 1} ---");
                    }

                    // Get all words on the page
                    IEnumerable<Word> words = page.GetWords();

                    if (orderByPosition)
                    {
                        // Order words by their position on the page (top to bottom, left to right)
                        words = words.OrderByDescending(w => w.BoundingBox.Top)
                                    .ThenBy(w => w.BoundingBox.Left);

                        // Group words by approximate line position
                        var lines = GroupWordsByLines(words.ToList());

                        // Output each line
                        foreach (var line in lines)
                        {
                            sb.AppendLine(string.Join(" ", line.Select(w => w.Text)));
                        }
                    }
                    else
                    {
                        // Simply extract text as is (may not preserve proper reading order)
                        sb.AppendLine(string.Join(" ", words.Select(w => w.Text)));
                    }

                    // Add page break if requested and not on the last page
                    if (preservePageBreaks && i < document.NumberOfPages - 1)
                    {
                        sb.AppendLine();
                        sb.AppendLine("==========================================");
                        sb.AppendLine();
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Groups words by lines based on their Y position.
        /// </summary>
        /// <param name="words">The list of words to group.</param>
        /// <returns>A list of lines, where each line is a list of words.</returns>
        private List<List<Word>> GroupWordsByLines(List<Word> words)
        {
            var lines = new List<List<Word>>();
            if (words.Count == 0)
                return lines;

            const double lineHeightThreshold = 5.0; // Adjust based on your documents

            // Sort words by Y position (from top to bottom)
            var sortedWords = words.OrderByDescending(w => w.BoundingBox.Top).ToList();

            // Initialize the first line with the first word
            var currentLine = new List<Word> { sortedWords[0] };
            double currentY = sortedWords[0].BoundingBox.Top;

            // Group remaining words into lines
            for (int i = 1; i < sortedWords.Count; i++)
            {
                var word = sortedWords[i];
                double yDiff = Math.Abs(word.BoundingBox.Top - currentY);

                if (yDiff <= lineHeightThreshold)
                {
                    // Word is on the same line
                    currentLine.Add(word);
                }
                else
                {
                    // Sort words in the current line by X position (from left to right)
                    lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());

                    // Start a new line
                    currentLine = new List<Word> { word };
                    currentY = word.BoundingBox.Top;
                }
            }

            // Add the last line
            if (currentLine.Count > 0)
            {
                lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());
            }

            return lines;
        }
    }
}
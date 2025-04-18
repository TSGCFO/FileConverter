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
using DocumentFormat.OpenXml;

namespace FileConverter.Converters.Documents
{
    /// <summary>
    /// Converter implementation for extracting text content from DOCX files.
    /// </summary>
    public class DocxToTxtConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Docx };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Txt };

        /// <summary>
        /// Converts a DOCX file to text format by extracting its content.
        /// </summary>
        /// <param name="inputPath">Path to the input DOCX file.</param>
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
                    StatusMessage = "Starting DOCX to text conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                bool preserveLineBreaks = parameters.GetParameter("preserveLineBreaks", true);
                bool preserveHeadersFooters = parameters.GetParameter("preserveHeadersFooters", true);
                bool includeComments = parameters.GetParameter("includeComments", false);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading DOCX file..."
                });

                // Extract text from DOCX
                string extractedText = await Task.Run(() =>
                    ExtractTextFromDocx(inputPath, preserveLineBreaks, preserveHeadersFooters, includeComments),
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
                    InputFormat = FileFormat.Docx,
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
                    InputFormat = FileFormat.Docx,
                    OutputFormat = FileFormat.Txt,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts text content from a DOCX file.
        /// </summary>
        /// <param name="docxPath">Path to the DOCX file.</param>
        /// <param name="preserveLineBreaks">Whether to preserve paragraph breaks in the output.</param>
        /// <param name="preserveHeadersFooters">Whether to include headers and footers in the output.</param>
        /// <param name="includeComments">Whether to include document comments in the output.</param>
        /// <returns>The extracted text content.</returns>
        private string ExtractTextFromDocx(
            string docxPath,
            bool preserveLineBreaks,
            bool preserveHeadersFooters,
            bool includeComments)
        {
            var sb = new StringBuilder();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null || mainPart.Document == null || mainPart.Document.Body == null)
                {
                    return string.Empty;
                }

                // Extract text from the main document part
                ExtractTextFromPart(mainPart, sb, preserveLineBreaks);

                // Extract text from headers and footers if requested
                if (preserveHeadersFooters)
                {
                    sb.AppendLine();
                    sb.AppendLine("--- HEADERS ---");
                    foreach (var headerPart in mainPart.HeaderParts)
                    {
                        ExtractTextFromPart(headerPart, sb, preserveLineBreaks);
                    }

                    sb.AppendLine();
                    sb.AppendLine("--- FOOTERS ---");
                    foreach (var footerPart in mainPart.FooterParts)
                    {
                        ExtractTextFromPart(footerPart, sb, preserveLineBreaks);
                    }
                }

                // Extract comments if requested
                if (includeComments && mainPart.WordprocessingCommentsPart != null)
                {
                    sb.AppendLine();
                    sb.AppendLine("--- COMMENTS ---");

                    var comments = mainPart.WordprocessingCommentsPart.Comments;
                    if (comments != null)
                    {
                        foreach (var comment in comments.Elements<Comment>())
                        {
                            sb.AppendLine($"Comment by {comment.Author} ({comment.Date}):");
                            ExtractTextFromElement(comment, sb, preserveLineBreaks);
                            sb.AppendLine();
                        }
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Extracts text content from a document part.
        /// </summary>
        /// <param name="part">The document part to extract text from.</param>
        /// <param name="sb">The string builder to append text to.</param>
        /// <param name="preserveLineBreaks">Whether to preserve paragraph breaks in the output.</param>
        private void ExtractTextFromPart(OpenXmlPart part, StringBuilder sb, bool preserveLineBreaks)
        {
            if (part.RootElement == null)
                return;

            ExtractTextFromElement(part.RootElement, sb, preserveLineBreaks);
        }

        /// <summary>
        /// Extracts text content from an OpenXML element.
        /// </summary>
        /// <param name="element">The element to extract text from.</param>
        /// <param name="sb">The string builder to append text to.</param>
        /// <param name="preserveLineBreaks">Whether to preserve paragraph breaks in the output.</param>
        private void ExtractTextFromElement(OpenXmlElement element, StringBuilder sb, bool preserveLineBreaks)
        {
            // Handle specific paragraph cases
            if (element is Paragraph paragraph)
            {
                // Extract text from the paragraph
                var paragraphText = string.Join(" ", paragraph.Descendants<Text>().Select(t => t.Text));

                // Add the paragraph text if it's not empty
                if (!string.IsNullOrWhiteSpace(paragraphText))
                {
                    sb.Append(paragraphText);

                    // Add line break after paragraphs if requested
                    if (preserveLineBreaks)
                    {
                        sb.AppendLine();
                    }
                    else
                    {
                        sb.Append(" ");
                    }
                }
                // Add line break for empty paragraphs if preserving line breaks
                else if (preserveLineBreaks)
                {
                    sb.AppendLine();
                }
            }
            // Handle tables specially
            else if (element is Table table)
            {
                foreach (var row in table.Elements<TableRow>())
                {
                    // Extract text from each cell in the row
                    var cellTexts = row.Elements<TableCell>()
                        .Select(cell => string.Join(" ", cell.Descendants<Text>().Select(t => t.Text)))
                        .Where(text => !string.IsNullOrWhiteSpace(text));

                    // Join cell texts with tabs
                    string rowText = string.Join("\t", cellTexts);
                    if (!string.IsNullOrWhiteSpace(rowText))
                    {
                        sb.AppendLine(rowText);
                    }
                }
                sb.AppendLine();
            }
            else
            {
                // Recursively process child elements
                foreach (var child in element.Elements())
                {
                    ExtractTextFromElement(child, sb, preserveLineBreaks);
                }
            }
        }
    }
}
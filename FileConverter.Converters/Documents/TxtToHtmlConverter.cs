using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace FileConverter.Converters.Documents
{
    /// <summary>
    /// Converter implementation for converting plain text files to HTML format.
    /// </summary>
    public class TxtToHtmlConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Txt };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Html };

        /// <summary>
        /// Converts a text file to HTML format.
        /// </summary>
        /// <param name="inputPath">Path to the input text file.</param>
        /// <param name="outputPath">Path where the output HTML file will be saved.</param>
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
                    StatusMessage = "Starting text to HTML conversion..."
                });

                // Validate input and output
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Read the text file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading text file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                string content = await File.ReadAllTextAsync(inputPath, cancellationToken);

                // Convert the content to HTML
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 50,
                    StatusMessage = "Converting to HTML format..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                string htmlContent = ConvertTextToHtml(content, parameters);

                // Write the HTML file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing HTML file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                await File.WriteAllTextAsync(outputPath, htmlContent, Encoding.UTF8, cancellationToken);

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
                    OutputFormat = FileFormat.Html,
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
                    OutputFormat = FileFormat.Html,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Converts plain text content to HTML.
        /// </summary>
        /// <param name="textContent">The plain text content to convert.</param>
        /// <param name="parameters">Optional parameters for customizing the conversion.</param>
        /// <returns>The HTML representation of the text content.</returns>
        private string ConvertTextToHtml(string textContent, ConversionParameters parameters)
        {
            // Get custom parameters or use defaults
            string title = parameters.GetParameter("title", "Converted Document");
            string cssStyle = parameters.GetParameter("css", DefaultCss);
            bool preserveLineBreaks = parameters.GetParameter("preserveLineBreaks", true);

            // Escape the text content to prevent HTML injection
            string escapedContent = HttpUtility.HtmlEncode(textContent);

            // Process line breaks if needed
            if (preserveLineBreaks)
            {
                escapedContent = escapedContent.Replace("\r\n", "<br>")
                                            .Replace("\n", "<br>")
                                            .Replace("\r", "<br>");
            }

            // Create the HTML document
            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.AppendLine("<!DOCTYPE html>");
            htmlBuilder.AppendLine("<html lang=\"en\">");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine($"  <meta charset=\"UTF-8\">");
            htmlBuilder.AppendLine($"  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            htmlBuilder.AppendLine($"  <title>{HttpUtility.HtmlEncode(title)}</title>");
            htmlBuilder.AppendLine($"  <style>");
            htmlBuilder.AppendLine($"    {cssStyle}");
            htmlBuilder.AppendLine($"  </style>");
            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body>");
            htmlBuilder.AppendLine($"  <div class=\"content\">");
            htmlBuilder.AppendLine($"    {escapedContent}");
            htmlBuilder.AppendLine($"  </div>");
            htmlBuilder.AppendLine("</body>");
            htmlBuilder.AppendLine("</html>");

            return htmlBuilder.ToString();
        }

        /// <summary>
        /// Default CSS styling for the HTML document.
        /// </summary>
        private const string DefaultCss = @"
            body {
                font-family: Arial, sans-serif;
                line-height: 1.6;
                margin: 0;
                padding: 20px;
                color: #333;
                max-width: 800px;
                margin: 0 auto;
            }
            .content {
                background-color: #fff;
                padding: 20px;
                border-radius: 5px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            }
        ";
    }
}
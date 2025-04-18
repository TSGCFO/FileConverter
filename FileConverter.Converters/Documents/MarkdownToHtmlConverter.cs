using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using Markdig;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Converters.Documents
{
    /// <summary>
    /// Converter implementation for converting Markdown files to HTML format.
    /// </summary>
    public class MarkdownToHtmlConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Md };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Html };

        /// <summary>
        /// Converts a Markdown file to HTML format.
        /// </summary>
        /// <param name="inputPath">Path to the input Markdown file.</param>
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
                    StatusMessage = "Starting Markdown to HTML conversion..."
                });

                // Validate input and output
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Read the Markdown file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading Markdown file..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                string markdownContent = await File.ReadAllTextAsync(inputPath, cancellationToken);

                // Convert the content to HTML
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 50,
                    StatusMessage = "Converting to HTML format..."
                });

                cancellationToken.ThrowIfCancellationRequested();

                string htmlContent = ConvertMarkdownToHtml(markdownContent, parameters);

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
                    InputFormat = FileFormat.Md,
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
                    InputFormat = FileFormat.Md,
                    OutputFormat = FileFormat.Html,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Converts Markdown content to HTML.
        /// </summary>
        /// <param name="markdownContent">The Markdown content to convert.</param>
        /// <param name="parameters">Optional parameters for customizing the conversion.</param>
        /// <returns>The HTML representation of the Markdown content.</returns>
        private string ConvertMarkdownToHtml(string markdownContent, ConversionParameters parameters)
        {
            // Get custom parameters or use defaults
            string title = parameters.GetParameter("title", "Converted Document");
            string cssStyle = parameters.GetParameter("css", DefaultCss);
            bool useAdvancedExtensions = parameters.GetParameter("useAdvancedExtensions", true);

            // Configure Markdown pipeline
            var pipelineBuilder = new MarkdownPipelineBuilder();

            if (useAdvancedExtensions)
            {
                // Enable GitHub-flavored Markdown and other useful extensions
                pipelineBuilder.UseAdvancedExtensions();
            }

            var pipeline = pipelineBuilder.Build();

            // Convert Markdown to HTML
            string htmlBody = Markdown.ToHtml(markdownContent, pipeline);

            // Create a complete HTML document
            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.AppendLine("<!DOCTYPE html>");
            htmlBuilder.AppendLine("<html lang=\"en\">");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine($"  <meta charset=\"UTF-8\">");
            htmlBuilder.AppendLine($"  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            htmlBuilder.AppendLine($"  <title>{System.Web.HttpUtility.HtmlEncode(title)}</title>");
            htmlBuilder.AppendLine($"  <style>");
            htmlBuilder.AppendLine($"    {cssStyle}");
            htmlBuilder.AppendLine($"  </style>");
            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body>");
            htmlBuilder.AppendLine($"  <div class=\"content\">");
            htmlBuilder.AppendLine($"    {htmlBody}");
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
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
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
            h1, h2, h3, h4, h5, h6 {
                margin-top: 1.5em;
                margin-bottom: 0.5em;
                font-weight: 600;
            }
            h1 { font-size: 2em; }
            h2 { font-size: 1.6em; }
            h3 { font-size: 1.4em; }
            h4 { font-size: 1.2em; }
            h5 { font-size: 1.1em; }
            h6 { font-size: 1em; }
            code {
                background-color: #f5f5f5;
                padding: 0.2em 0.4em;
                border-radius: 3px;
                font-family: Consolas, Monaco, 'Andale Mono', 'Ubuntu Mono', monospace;
                font-size: 85%;
            }
            pre {
                background-color: #f5f5f5;
                padding: 1em;
                border-radius: 5px;
                overflow-x: auto;
            }
            pre code {
                background-color: transparent;
                padding: 0;
            }
            blockquote {
                border-left: 4px solid #ddd;
                padding-left: 1em;
                margin-left: 0;
                color: #666;
            }
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 1em 0;
            }
            table, th, td {
                border: 1px solid #ddd;
            }
            th, td {
                padding: 0.5em;
                text-align: left;
            }
            th {
                background-color: #f5f5f5;
            }
            img {
                max-width: 100%;
                height: auto;
            }
        ";
    }
}
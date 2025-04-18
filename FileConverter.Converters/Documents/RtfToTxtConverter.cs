using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Converters.Documents
{
    /// <summary>
    /// Converter implementation for extracting text content from RTF files.
    /// </summary>
    public class RtfToTxtConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Rtf };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Txt };

        /// <summary>
        /// Converts an RTF file to text format by extracting its content.
        /// </summary>
        /// <param name="inputPath">Path to the input RTF file.</param>
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
                    StatusMessage = "Starting RTF to text conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                bool preserveLineBreaks = parameters.GetParameter("preserveLineBreaks", true);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading RTF file..."
                });

                // Extract text from RTF
                string extractedText = await Task.Run(() =>
                    ExtractTextFromRtf(inputPath, preserveLineBreaks),
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
                    InputFormat = FileFormat.Rtf,
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
                    InputFormat = FileFormat.Rtf,
                    OutputFormat = FileFormat.Txt,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts text content from an RTF file.
        /// </summary>
        /// <param name="rtfPath">Path to the RTF file.</param>
        /// <param name="preserveLineBreaks">Whether to preserve paragraph breaks in the output.</param>
        /// <returns>The extracted text content.</returns>
        private string ExtractTextFromRtf(string rtfPath, bool preserveLineBreaks)
        {
            // Read the RTF file as raw text
            string rtfContent = File.ReadAllText(rtfPath);
            string plainText = StripRtfTags(rtfContent);

            // Process line breaks if needed
            if (!preserveLineBreaks)
            {
                // Replace line breaks with spaces
                plainText = plainText.Replace("\n", " ").Replace("\r", " ");

                // Normalize whitespace
                while (plainText.Contains("  "))
                {
                    plainText = plainText.Replace("  ", " ");
                }
            }

            return plainText;
        }

        /// <summary>
        /// Strips RTF tags from the input string to extract plain text.
        /// </summary>
        /// <param name="rtfText">The RTF text to strip.</param>
        /// <returns>Plain text without RTF tags.</returns>
        private string StripRtfTags(string rtfText)
        {
            // Check if it's a valid RTF document
            if (!rtfText.StartsWith("{\\rtf"))
            {
                return rtfText; // Not an RTF document, return as is
            }

            StringBuilder result = new StringBuilder();
            bool inControlWord = false;
            bool inGroup = false;
            int bracketCount = 0;

            // Process character by character
            for (int i = 0; i < rtfText.Length; i++)
            {
                char c = rtfText[i];

                if (c == '{')
                {
                    inGroup = true;
                    bracketCount++;
                    continue;
                }

                if (c == '}')
                {
                    bracketCount--;
                    if (bracketCount <= 0)
                    {
                        inGroup = false;
                    }
                    continue;
                }

                if (c == '\\')
                {
                    inControlWord = true;

                    // Check for special characters
                    if (i + 1 < rtfText.Length)
                    {
                        char nextChar = rtfText[i + 1];

                        // Handle escaped characters like \', \", etc.
                        if (nextChar == '\'')
                        {
                            // This is a hex-encoded character (e.g. \'a9 for ©)
                            if (i + 3 < rtfText.Length && char.IsLetterOrDigit(rtfText[i + 2]) && char.IsLetterOrDigit(rtfText[i + 3]))
                            {
                                string hexStr = rtfText.Substring(i + 2, 2);
                                try
                                {
                                    int charCode = Convert.ToInt32(hexStr, 16);
                                    result.Append((char)charCode);
                                }
                                catch
                                {
                                    // Ignore if we can't convert
                                }

                                i += 3; // Skip \'xx
                                inControlWord = false;
                            }
                        }
                        else if (nextChar == '\\' || nextChar == '{' || nextChar == '}')
                        {
                            // These are escaped characters, add them
                            result.Append(nextChar);
                            i++; // Skip to the next character
                            inControlWord = false;
                        }
                    }

                    continue;
                }

                if (inControlWord)
                {
                    // Skip control word until we reach a space or non-letter
                    if (char.IsLetter(c))
                    {
                        continue;
                    }

                    // Skip the parameter (if any) after control word
                    if (char.IsDigit(c) || c == '-')
                    {
                        continue;
                    }

                    inControlWord = false;

                    // If there's a space after the control word, skip it too
                    if (c == ' ')
                    {
                        continue;
                    }
                }

                // Process special whitespace and line break tokens
                if (rtfText.Substring(i).StartsWith("\\par") || rtfText.Substring(i).StartsWith("\\line"))
                {
                    result.AppendLine();
                    i += 4; // Skip \par or \line
                    continue;
                }

                // Only add the character if we're not in a control word
                if (!inControlWord && !inGroup)
                {
                    result.Append(c);
                }
            }

            return result.ToString().Trim();
        }
    }
}
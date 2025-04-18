﻿using FileConverter.Common.Enums;
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
    /// Converter implementation for converting RTF files to TSV format.
    /// </summary>
    public class RtfToTsvConverter : IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedInputFormats => new[] { FileFormat.Rtf };

        /// <summary>
        /// Gets the supported output formats for this converter.
        /// </summary>
        public IEnumerable<FileFormat> SupportedOutputFormats => new[] { FileFormat.Tsv };

        /// <summary>
        /// Converts an RTF file to TSV format.
        /// </summary>
        /// <param name="inputPath">Path to the input RTF file.</param>
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
                    StatusMessage = "Starting RTF to TSV conversion..."
                });

                // Validate input
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException("Input file not found", inputPath);
                }

                // Get parameters
                string lineDelimiter = parameters.GetParameter("lineDelimiter", string.Empty);
                bool treatFirstLineAsHeader = parameters.GetParameter("treatFirstLineAsHeader", false);

                // Report reading progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 20,
                    StatusMessage = "Reading RTF file..."
                });

                // Extract text from RTF
                string extractedText = await Task.Run(() =>
                    ExtractTextFromRtf(inputPath),
                    cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Convert to TSV
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 50,
                    StatusMessage = "Converting to TSV format..."
                });

                string tsvContent = await Task.Run(() =>
                    ConvertToTsv(extractedText, lineDelimiter, treatFirstLineAsHeader),
                    cancellationToken);

                cancellationToken.ThrowIfCancellationRequested();

                // Write the TSV file
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 80,
                    StatusMessage = "Writing TSV file..."
                });

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
                    InputFormat = FileFormat.Rtf,
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
                    InputFormat = FileFormat.Rtf,
                    OutputFormat = FileFormat.Tsv,
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Extracts text content from an RTF file.
        /// </summary>
        /// <param name="rtfPath">Path to the RTF file.</param>
        /// <returns>The extracted text content.</returns>
        private string ExtractTextFromRtf(string rtfPath)
        {
            // Read the RTF file as raw text
            string rtfContent = File.ReadAllText(rtfPath);
            return StripRtfTags(rtfContent);
        }

        /// <summary>
        /// Converts plain text to TSV format.
        /// </summary>
        /// <param name="text">The text to convert.</param>
        /// <param name="lineDelimiter">Character or string used to split each line into columns.</param>
        /// <param name="treatFirstLineAsHeader">Whether to treat the first line as a header.</param>
        /// <returns>The TSV content.</returns>
        private string ConvertToTsv(string text, string lineDelimiter, bool treatFirstLineAsHeader)
        {
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var tsvBuilder = new StringBuilder();

            int startLine = 0;

            // Process all lines
            for (int i = startLine; i < lines.Length; i++)
            {
                // Skip empty lines
                if (string.IsNullOrWhiteSpace(lines[i]))
                    continue;

                string line = lines[i].Trim();

                // Convert the line to a TSV row
                if (!string.IsNullOrEmpty(lineDelimiter))
                {
                    // Split the line using the delimiter and create a TSV row
                    string[] fields = line.Split(lineDelimiter);
                    string tsvLine = string.Join("\t", fields.Select(field => EscapeForTsv(field.Trim())));
                    tsvBuilder.AppendLine(tsvLine);
                }
                else
                {
                    // No delimiter, treat the whole line as a single field
                    tsvBuilder.AppendLine(EscapeForTsv(line));
                }
            }

            return tsvBuilder.ToString();
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
                if (i + 3 < rtfText.Length &&
                    ((rtfText[i] == '\\' && rtfText[i + 1] == 'p' && rtfText[i + 2] == 'a' && rtfText[i + 3] == 'r') ||
                     (rtfText[i] == '\\' && rtfText[i + 1] == 'l' && rtfText[i + 2] == 'i' && rtfText[i + 3] == 'n')))
                {
                    result.AppendLine();
                    i += 3; // Skip par or lin (from \par or \line)
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
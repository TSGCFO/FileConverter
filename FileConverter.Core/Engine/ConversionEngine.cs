using FileConverter.Common.Enums;
using FileConverter.Common.Interfaces;
using FileConverter.Common.Models;
using FileConverter.Core.Engine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Core.Engine
{
    /// <summary>
    /// Main engine responsible for coordinating file conversions.
    /// </summary>
    public class ConversionEngine
    {
        private readonly IEnumerable<IConverter> _converters;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversionEngine"/> class.
        /// </summary>
        /// <param name="converters">The collection of available converters.</param>
        public ConversionEngine(IEnumerable<IConverter> converters)
        {
            _converters = converters ?? throw new ArgumentNullException(nameof(converters));
        }

        /// <summary>
        /// Gets all available converters.
        /// </summary>
        public IEnumerable<IConverter> Converters => _converters;

        /// <summary>
        /// Converts a file from one format to another.
        /// </summary>
        /// <param name="inputPath">The path to the input file.</param>
        /// <param name="outputPath">The path where the output file will be saved.</param>
        /// <param name="parameters">Optional parameters for customizing the conversion.</param>
        /// <param name="progress">An optional interface for reporting progress.</param>
        /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
        /// <returns>A conversion result with details about the operation.</returns>
        public async Task<ConversionResult> ConvertFileAsync(
            string inputPath,
            string outputPath,
            ConversionParameters? parameters = null,
            IProgress<ConversionProgress>? progress = null,
            CancellationToken cancellationToken = default)
        {
            // Validate parameters
            if (string.IsNullOrEmpty(inputPath))
                throw new ArgumentException("Input path cannot be empty", nameof(inputPath));

            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path cannot be empty", nameof(outputPath));

            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input file not found", inputPath);

            parameters ??= new ConversionParameters();

            var startTime = DateTime.Now;

            try
            {
                // Detect formats
                var inputFormat = FormatDetector.DetectFormat(inputPath);
                var outputFormat = FormatDetector.DetectFormat(outputPath);

                if (inputFormat == FileFormat.Unknown)
                    throw new InvalidOperationException($"Unknown input format: {Path.GetExtension(inputPath)}");

                if (outputFormat == FileFormat.Unknown)
                    throw new InvalidOperationException($"Unknown output format: {Path.GetExtension(outputPath)}");

                // Find suitable converter
                var converter = FindConverter(inputFormat, outputFormat);
                if (converter == null)
                    throw new InvalidOperationException(
                        $"No converter found for {inputFormat} to {outputFormat} conversion");

                // Report initial progress
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 0,
                    StatusMessage = "Starting conversion..."
                });

                // Ensure output directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? string.Empty);

                // Execute conversion
                var result = await converter.ConvertAsync(
                    inputPath,
                    outputPath,
                    parameters,
                    progress,
                    cancellationToken);

                result.ElapsedTime = DateTime.Now - startTime;
                return result;
            }
            catch (OperationCanceledException)
            {
                progress?.Report(new ConversionProgress
                {
                    PercentComplete = 0,
                    StatusMessage = "Operation canceled."
                });

                return new ConversionResult
                {
                    Success = false,
                    InputPath = inputPath,
                    OutputPath = outputPath,
                    InputFormat = FormatDetector.DetectFormat(inputPath),
                    OutputFormat = FormatDetector.DetectFormat(outputPath),
                    ElapsedTime = DateTime.Now - startTime,
                    Error = new OperationCanceledException("Conversion was canceled")
                };
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
                    InputFormat = FormatDetector.DetectFormat(inputPath),
                    OutputFormat = FormatDetector.DetectFormat(outputPath),
                    ElapsedTime = DateTime.Now - startTime,
                    Error = ex
                };
            }
        }

        /// <summary>
        /// Finds a converter that can convert from the specified input format to the specified output format.
        /// </summary>
        /// <param name="inputFormat">The input file format.</param>
        /// <param name="outputFormat">The output file format.</param>
        /// <returns>A converter that can perform the conversion, or null if none is found.</returns>
        private IConverter? FindConverter(FileFormat inputFormat, FileFormat outputFormat)
        {
            return _converters.FirstOrDefault(c =>
                c.SupportedInputFormats.Contains(inputFormat) &&
                c.SupportedOutputFormats.Contains(outputFormat));
        }

        /// <summary>
        /// Gets all supported conversion paths (input format to output format).
        /// </summary>
        /// <returns>A collection of tuples representing the supported conversion paths.</returns>
        public IEnumerable<(FileFormat InputFormat, FileFormat OutputFormat)> GetSupportedConversionPaths()
        {
            var paths = new List<(FileFormat, FileFormat)>();

            foreach (var converter in _converters)
            {
                foreach (var inputFormat in converter.SupportedInputFormats)
                {
                    foreach (var outputFormat in converter.SupportedOutputFormats)
                    {
                        paths.Add((inputFormat, outputFormat));
                    }
                }
            }

            return paths;
        }
    }
}
using FileConverter.Common.Enums;
using FileConverter.Common.Models;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace FileConverter.Common.Interfaces
{
    /// <summary>
    /// Interface for file format converters
    /// </summary>
    public interface IConverter
    {
        /// <summary>
        /// Gets the supported input formats for this converter
        /// </summary>
        IEnumerable<FileFormat> SupportedInputFormats { get; }

        /// <summary>
        /// Gets the supported output formats for this converter
        /// </summary>
        IEnumerable<FileFormat> SupportedOutputFormats { get; }

        /// <summary>
        /// Converts a file from one format to another
        /// </summary>
        /// <param name="inputPath">Path to the input file</param>
        /// <param name="outputPath">Path where the output file will be saved</param>
        /// <param name="parameters">Conversion parameters</param>
        /// <param name="progress">Progress reporting interface</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>Conversion result with details about the operation</returns>
        Task<ConversionResult> ConvertAsync(
            string inputPath,
            string outputPath,
            ConversionParameters parameters,
            IProgress<ConversionProgress>? progress,
            CancellationToken cancellationToken);
    }
}
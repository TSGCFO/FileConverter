using FileConverter.Common.Enums;
using System;

namespace FileConverter.Common.Models
{
    /// <summary>
    /// Represents the result of a file conversion operation
    /// </summary>
    public class ConversionResult
    {
        /// <summary>
        /// Gets or sets a value indicating whether the conversion was successful
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Gets or sets the path to the input file
        /// </summary>
        public string InputPath { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the path to the output file
        /// </summary>
        public string OutputPath { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the format of the input file
        /// </summary>
        public FileFormat InputFormat { get; set; }

        /// <summary>
        /// Gets or sets the format of the output file
        /// </summary>
        public FileFormat OutputFormat { get; set; }

        /// <summary>
        /// Gets or sets the time taken to complete the conversion
        /// </summary>
        public TimeSpan ElapsedTime { get; set; }

        /// <summary>
        /// Gets or sets the error that occurred during conversion, if any
        /// </summary>
        public Exception? Error { get; set; }
    }
}
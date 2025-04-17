using FileConverter.Common.Models;
using FileConverter.Core.Logging;
using System;

namespace FileConverter.Core.Helpers
{
    /// <summary>
    /// Helper class for reporting progress during conversion operations.
    /// </summary>
    public class ProgressReporter : IProgress<ConversionProgress>
    {
        private readonly IProgress<ConversionProgress>? _innerProgress;
        private readonly string _inputPath;
        private bool _logProgress;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProgressReporter"/> class.
        /// </summary>
        /// <param name="inputPath">The input file path (used for logging).</param>
        /// <param name="innerProgress">The inner progress reporter to forward reports to.</param>
        /// <param name="logProgress">Whether to log progress updates.</param>
        public ProgressReporter(
            string inputPath,
            IProgress<ConversionProgress>? innerProgress = null,
            bool logProgress = true)
        {
            _inputPath = inputPath;
            _innerProgress = innerProgress;
            _logProgress = logProgress;
        }

        /// <summary>
        /// Reports a progress update.
        /// </summary>
        /// <param name="value">The current progress value.</param>
        public void Report(ConversionProgress value)
        {
            // Forward to the inner progress reporter if available
            _innerProgress?.Report(value);

            // Log the progress if enabled
            if (_logProgress)
            {
                ConversionLogger.LogProgress(value, _inputPath);
            }
        }

        /// <summary>
        /// Creates a new progress report with the specified percentage and message.
        /// </summary>
        /// <param name="percentComplete">The percentage complete (0-100).</param>
        /// <param name="statusMessage">The status message.</param>
        public void ReportProgress(float percentComplete, string statusMessage)
        {
            Report(new ConversionProgress
            {
                PercentComplete = percentComplete,
                StatusMessage = statusMessage
            });
        }
    }
}
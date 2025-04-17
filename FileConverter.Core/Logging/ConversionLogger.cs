using FileConverter.Common.Models;
using System;
using System.IO;

namespace FileConverter.Core.Logging
{
    /// <summary>
    /// Provides logging capabilities for conversion operations.
    /// </summary>
    public static class ConversionLogger
    {
        private static string _logFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FileConverter",
            "logs",
            "conversion.log");

        /// <summary>
        /// Gets or sets the path to the log file.
        /// </summary>
        public static string LogFilePath
        {
            get => _logFilePath;
            set
            {
                _logFilePath = value;
                // Ensure the directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath) ?? string.Empty);
            }
        }

        /// <summary>
        /// Logs the start of a conversion operation.
        /// </summary>
        /// <param name="inputPath">The input file path.</param>
        /// <param name="outputPath">The output file path.</param>
        public static void LogConversionStart(string inputPath, string outputPath)
        {
            LogMessage($"Starting conversion: {inputPath} -> {outputPath}");
        }

        /// <summary>
        /// Logs the completion of a conversion operation.
        /// </summary>
        /// <param name="result">The conversion result.</param>
        public static void LogConversionComplete(ConversionResult result)
        {
            string status = result.Success ? "SUCCESS" : "FAILED";
            string errorMessage = result.Error != null ? $" - Error: {result.Error.Message}" : string.Empty;

            LogMessage($"Conversion completed [{status}]: {result.InputPath} -> {result.OutputPath} " +
                       $"({result.ElapsedTime.TotalSeconds:F2} seconds){errorMessage}");
        }

        /// <summary>
        /// Logs a progress update during conversion.
        /// </summary>
        /// <param name="progress">The conversion progress.</param>
        /// <param name="inputPath">The input file path.</param>
        public static void LogProgress(ConversionProgress progress, string inputPath)
        {
            LogMessage($"Progress ({Path.GetFileName(inputPath)}): {progress.PercentComplete:F0}% - {progress.StatusMessage}");
        }

        /// <summary>
        /// Logs a general message to the log file.
        /// </summary>
        /// <param name="message">The message to log.</param>
        public static void LogMessage(string message)
        {
            try
            {
                // Ensure the directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath) ?? string.Empty);

                // Write the log entry
                using (StreamWriter writer = File.AppendText(LogFilePath))
                {
                    writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
                }
            }
            catch (Exception ex)
            {
                // If logging fails, we don't want to crash the application
                Console.Error.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }
    }
}
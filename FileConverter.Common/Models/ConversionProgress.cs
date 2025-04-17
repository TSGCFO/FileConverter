namespace FileConverter.Common.Models
{
    /// <summary>
    /// Represents the progress of a conversion operation
    /// </summary>
    public class ConversionProgress
    {
        /// <summary>
        /// Gets or sets the percentage of completion (0-100)
        /// </summary>
        public float PercentComplete { get; set; }

        /// <summary>
        /// Gets or sets a descriptive message about the current status
        /// </summary>
        public string StatusMessage { get; set; } = string.Empty;
    }
}
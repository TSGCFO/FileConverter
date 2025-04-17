using System.Collections.Generic;

namespace FileConverter.Common.Models
{
    /// <summary>
    /// Represents parameters for customizing a conversion operation
    /// </summary>
    public class ConversionParameters
    {
        /// <summary>
        /// Gets or sets a dictionary of named parameters and their values
        /// </summary>
        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// Adds a parameter with the specified name and value
        /// </summary>
        public void AddParameter(string name, object value)
        {
            Parameters[name] = value;
        }

        /// <summary>
        /// Gets a parameter value by name, or returns the default value if not found
        /// </summary>
        public T GetParameter<T>(string name, T defaultValue = default)
        {
            if (Parameters.TryGetValue(name, out object? value) && value is T typedValue)
            {
                return typedValue;
            }

            return defaultValue;
        }
    }
}
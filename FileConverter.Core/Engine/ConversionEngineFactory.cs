using FileConverter.Common.Interfaces;
using System.Collections.Generic;
using FileConverter.Converters.Documents;
using FileConverter.Converters.Spreadsheets;

namespace FileConverter.Core.Engine
{
    /// <summary>
    /// Factory for creating instances of the conversion engine with the appropriate converters.
    /// </summary>
    public static class ConversionEngineFactory
    {
        /// <summary>
        /// Creates a conversion engine with the specified converters.
        /// </summary>
        /// <param name="converters">The converters to use.</param>
        /// <returns>A configured conversion engine.</returns>
        public static ConversionEngine Create(IEnumerable<IConverter> converters)
        {
            return new ConversionEngine(converters);
        }

        /// <summary>
        /// Creates a conversion engine with the default set of converters.
        /// </summary>
        /// <returns>A configured conversion engine.</returns>
        public static ConversionEngine CreateWithDefaultConverters()
        {
            var converters = new List<IConverter>
            {
                new TxtToHtmlConverter(),
                new TsvToCsvConverter(),
                new CsvToTsvConverter(),
                new HtmlToTsvConverter(),
                new MarkdownToTsvConverter(),
                new TxtToTsvConverter(),
                new TxtToCsvConverter(),
                new JsonToCsvConverter(),
                new XmlToCsvConverter(),
                new XmlToTsvConverter(),
                new JsonToTsvConverter(),
                // Add the new converters here
                new DocxToCsvConverter(),
                new DocxToTsvConverter(),
                new PdfToCsvConverter(),
                new PdfToTsvConverter(),
                new MarkdownToHtmlConverter(),
                new PdfToTxtConverter(),
                // Add more converters here as they are implemented
            };

            return new ConversionEngine(converters);
        }
    }
}
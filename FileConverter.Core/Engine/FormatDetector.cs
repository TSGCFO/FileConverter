using FileConverter.Common.Enums;
using System;
using System.IO;

namespace FileConverter.Core.Engine
{
    /// <summary>
    /// Provides methods to detect file formats based on file extensions.
    /// </summary>
    public class FormatDetector
    {
        /// <summary>
        /// Detects the file format based on the file extension.
        /// </summary>
        /// <param name="filePath">The path to the file.</param>
        /// <returns>The detected file format, or FileFormat.Unknown if the format cannot be determined.</returns>
        public static FileFormat DetectFormat(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return FileFormat.Unknown;

            // Get the file extension (lowercase, without the dot)
            string extension = Path.GetExtension(filePath).ToLowerInvariant().TrimStart('.');

            // Map the extension to a FileFormat
            return extension switch
            {
                // Document formats
                "pdf" => FileFormat.Pdf,
                "doc" => FileFormat.Doc,
                "docx" => FileFormat.Docx,
                "rtf" => FileFormat.Rtf,
                "odt" => FileFormat.Odt,
                "txt" => FileFormat.Txt,
                "html" or "htm" => FileFormat.Html,
                "md" or "markdown" => FileFormat.Md,

                // Spreadsheet formats
                "xlsx" => FileFormat.Xlsx,
                "xls" => FileFormat.Xls,
                "csv" => FileFormat.Csv,
                "tsv" => FileFormat.Tsv,

                // Image formats
                "jpg" or "jpeg" => FileFormat.Jpeg,
                "png" => FileFormat.Png,
                "bmp" => FileFormat.Bmp,
                "gif" => FileFormat.Gif,
                "tiff" or "tif" => FileFormat.Tiff,
                "webp" => FileFormat.Webp,

                // Data formats
                "json" => FileFormat.Json,
                "xml" => FileFormat.Xml,
                "yaml" or "yml" => FileFormat.Yaml,
                "ini" => FileFormat.Ini,
                "toml" => FileFormat.Toml,

                // Archive formats
                "zip" => FileFormat.Zip,
                "tar" => FileFormat.Tar,
                "gz" or "gzip" => FileFormat.Gz,
                "7z" => FileFormat.SevenZ,

                // Unknown format
                _ => FileFormat.Unknown
            };
        }
    }
}
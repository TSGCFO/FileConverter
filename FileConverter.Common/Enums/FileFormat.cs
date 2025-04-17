namespace FileConverter.Common.Enums
{
    /// <summary>
    /// Defines all supported file formats for conversion
    /// </summary>
    public enum FileFormat
    {
        Unknown = 0,

        // Document formats
        Pdf = 1,
        Doc = 2,
        Docx = 3,
        Rtf = 4,
        Odt = 5,
        Txt = 6,
        Html = 7,
        Md = 8,

        // Spreadsheet formats
        Xlsx = 20,
        Xls = 21,
        Csv = 22,
        Tsv = 23,

        // Image formats
        Jpeg = 40,
        Png = 41,
        Bmp = 42,
        Gif = 43,
        Tiff = 44,
        Webp = 45,

        // Data formats
        Json = 60,
        Xml = 61,
        Yaml = 62,
        Ini = 63,
        Toml = 64,

        // Archive formats
        Zip = 80,
        Tar = 81,
        Gz = 82,
        SevenZ = 83
    }
}
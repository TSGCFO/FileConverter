using FileConverter.Common.Models;
using FileConverter.Core.Engine;
using System;
using System.IO;
using System.Threading.Tasks;

namespace FileConverter.CLI
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("FileConverter CLI");
            Console.WriteLine("================");

            // Create a test text file
            // Create a test text file in the current directory
            string currentDirectory = Directory.GetCurrentDirectory();
            string testFilePath = Path.Combine(currentDirectory, "test.txt");
            string outputFilePath = Path.Combine(currentDirectory, "test.html");

            Console.WriteLine($"Creating test file: {testFilePath}");
            File.WriteAllText(testFilePath, "Hello, world!\n\nThis is a test file for the FileConverter application.\nIt demonstrates converting a text file to HTML.");

            // Create the conversion engine with default converters
            var engine = ConversionEngineFactory.CreateWithDefaultConverters();

            // Create a progress reporter
            var progress = new Progress<ConversionProgress>(p =>
            {
                Console.WriteLine($"Progress: {p.PercentComplete}% - {p.StatusMessage}");
            });

            // Perform the conversion
            Console.WriteLine($"Converting {testFilePath} to {outputFilePath}...");

            try
            {
                var result = await engine.ConvertFileAsync(testFilePath, outputFilePath, progress: progress);

                if (result.Success)
                {
                    Console.WriteLine($"Conversion successful! Time: {result.ElapsedTime.TotalSeconds:F2} seconds");
                    Console.WriteLine($"Output file: {Path.GetFullPath(outputFilePath)}");
                }
                else
                {
                    Console.WriteLine($"Conversion failed: {result.Error?.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}
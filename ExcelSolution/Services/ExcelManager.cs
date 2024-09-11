using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Runtime.CompilerServices;

namespace ExcelSolution.Services
{
    /// <summary>
    /// The ExcelManager class provides helper methods for working with Excel files.
    /// This includes file existence checks, retrieving the file path, and examining worksheets.
    /// </summary>
    public class ExcelManager
    {
        /// <summary>
        /// Checks whether the specified Excel file exists and can be read.
        /// </summary>
        /// <param name="fileName">The name of the Excel file to check.</param>
        /// <returns>True if the file exists and can be opened, otherwise false.</returns>
        protected bool FileExist(string fileName)
        {
            string excelFilePath = ExcelFilePath(fileName);

            // Check if the file exists at the given path.
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("Error: The Excel file does not exist.");
                return false;
            }

            // Try opening the file to check if it can be read.
            try
            {
                using (var stream = File.OpenRead(excelFilePath))
                {
                    Console.WriteLine("The Excel file exists and can be read.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: The Excel file could not be opened. Details: {ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Constructs and returns the full path to the specified Excel file.
        /// </summary>
        /// <param name="fileName">The name of the Excel file.</param>
        /// <returns>The full file path as a string.</returns>
        protected string ExcelFilePath(string fileName)
        {
            var rootDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;
            var excelFilePath = Path.Combine(rootDirectory, "ExcelSolution", "ExcelFiles", fileName);
            return excelFilePath;
        }

        /// <summary>
        /// Examines whether the specified worksheet exists and is properly defined.
        /// </summary>
        /// <param name="package">The Excel package containing the workbook.</param>
        /// <param name="worksheet">The worksheet to examine (output parameter).</param>
        /// <param name="sheet">The index of the worksheet to examine.</param>
        /// <returns>True if the worksheet exists and is valid, otherwise false.</returns>
        protected bool ExaminExcelTable(ExcelPackage package, out ExcelWorksheet worksheet, int sheet)
        {
            worksheet = package.Workbook.Worksheets[sheet]; // Get the worksheet by index.

            // Check if the worksheet exists.
            if (worksheet == null)
            {
                Console.WriteLine($"Error: Worksheet {sheet} not found.");
                return false;
            }

            // Check if the worksheet has any data (dimension is not null).
            if (worksheet.Dimension == null)
            {
                Console.WriteLine("Error: Worksheet is empty or not properly defined.");
                return false;
            }

            return true; // The worksheet exists and has data.
        }
    }
}

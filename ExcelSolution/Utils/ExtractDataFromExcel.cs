using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using ExcelSolution.Objects;
using ExcelSolution.Services;

namespace ExcelTableExtraction.Utils
{
    /// <summary>
    /// This class demonstrates how to use the ExcelManager class to extract data from an Excel file (Example.xlsx).
    /// The Excel file should be located in the ExcelFiles folder under the project directory.
    /// </summary>

    public class ExtractDataFromExcel : ExcelManager
    {
        /// <summary>
        /// This method extracts data from the specified Excel file and returns a list of Users.
        /// It first checks if the file exists, then reads its contents and performs basic data manipulation.
        /// </summary>
     
        public List<Users> ExtractData()
        {
            // Create an empty list to hold the extracted user data.
            List<Users> emptyUsersObject = new List<Users>();
            string fileName = "Example.xlsx";  // Excel file name
            string excelFilePath = "";  // Path to the Excel file

            // Check if the Excel file exists using the inherited FileExist method from ExcelManager.
            if (FileExist(fileName))
            {
                // Retrieve the full path to the Excel file using the inherited ExcelFilePath method.
                excelFilePath = ExcelFilePath(fileName);
                ExcelWorksheet worksheet;

                // EPPlus requires setting the LicenseContext. (Here we set it to NonCommercial for personal or educational use).
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Load the Excel file using EPPlus.
                FileInfo fileInfo = new FileInfo(excelFilePath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Get the first worksheet in the Excel file (index 0).
                    worksheet = package.Workbook.Worksheets[0]; // Will check if sheet 0 (first sheet) exists.

                    // Use the inherited ExaminExcelTable method to check if the worksheet is valid.
                    if (ExaminExcelTable(package, out worksheet, 0) == false)
                    {
                        Console.WriteLine("Could not read data in sheet 0");
                        return emptyUsersObject;  // Return an empty list if the worksheet is invalid.
                    }

                    // Get the number of rows in the worksheet.
                    int rowCount = worksheet.Dimension.Rows;

                    // Extract the data from the Excel worksheet into a list of Users.
                    List<Users> extractedUsersFromExcel = dataExtraction(rowCount, worksheet);

                    // Example of data manipulation: converting the user's name to lowercase if it exists, or setting a default value if it's null.
                    foreach (Users userRow in extractedUsersFromExcel)
                    {
                        userRow.Name = userRow.Name?.ToLower() ?? "no name found";
                        userRow.Id = userRow.Id ?? "no id found";
                        userRow.Position = userRow.Position?.ToLower()?? "no position found";
                    }

                    // Return the list of extracted and manipulated users.
                    return extractedUsersFromExcel;
                }
            }

            // Return an empty list if the file doesn't exist or couldn't be processed.
            return emptyUsersObject;
        }

        // This method extracts user data from the Excel worksheet and returns it as a list of Users.
        // It assumes the data starts at row 2 (row 1 contains headers).
        private List<Users> dataExtraction(int count, ExcelWorksheet worksheet)
        {
            List<Users> users = new List<Users>();

            // Loop through the rows in the worksheet, starting from row 2 (assuming the first row contains headers).
            for (int row = 2; row <= count; row++)
            {
                // Create a new Users object and populate its properties from the Excel columns.
                Users user = new Users
                {
                    Name = worksheet.Cells[row, 1].Text,
                    Id = worksheet.Cells[row, 2].Text,
                    Position = worksheet.Cells[row, 3].Text
                };

                // Add the user to the list.
                users.Add(user);
            }

            // Return the list of extracted users.
            return users;
        }
    }
}

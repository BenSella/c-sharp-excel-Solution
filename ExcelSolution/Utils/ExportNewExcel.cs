using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing.Text;
using System.ComponentModel;
using ExcelSolution.Objects;
using ExcelSolution.Services;

namespace ExcelTableExtraction.Utils
{
    public class ExportNewExcel : ExcelManager
    {
        public void ExportExcelFile(List<Users> users)
        {
            Random rand = new Random();
            string fileName = "ExampleForDataExport.xlsx";
            string excelFilePath = "";

            // Check if the file exists (can change file or method accordingly)
            if (FileExist(fileName))
            {
                excelFilePath = ExcelFilePath(fileName);
                ExcelWorksheet worksheet;

                // Set the license context explicitly (fully qualified to avoid ambiguity)
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    // Create a new worksheet named "Users"
                    worksheet = package.Workbook.Worksheets.Add("Users");
                    
                    foreach (Users user in users)
                    {
                        user.Seniority = Math.Round(rand.NextDouble() *(10.0-0.1)+0.1,1);// 1 decimal places
                    }
                    // Define the headers
                    worksheet.Cells[1, 1].Value = "Name";
                    worksheet.Cells[1, 2].Value = "ID";
                    worksheet.Cells[1, 3].Value = "User Position";
                    worksheet.Cells[1, 4].Value = "Seniority";

                    // Add user data to the worksheet
                    for (int i = 0; i < users.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = users[i].Name;
                        worksheet.Cells[i + 2, 2].Value = users[i].Id;
                        worksheet.Cells[i + 2, 3].Value = users[i].Position;
                        worksheet.Cells[i + 2, 4].Value = users[i].Seniority;
                    }

                    // Save the Excel package to the specified file path
                    FileInfo fileInfo = new FileInfo(excelFilePath);
                    package.SaveAs(fileInfo);
                    Console.WriteLine($"Excel file saved to: {excelFilePath}");
                }
            }
        }
    }
}

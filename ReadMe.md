# ExcelSolution: Excel Data Manipulation in C# with EPPlus nugget package

## Project Overview

**ExcelSolution** is a C# project demonstrating how to work with Excel files using the EPPlus 
library. This project allows users to extract data from Excel, manipulate it, and export 
the updated data back into a new Excel file. It provides a practical example of how to 
integrate file handling, object-oriented programming, and external libraries in a .NET application.

## Project Architecture

```
ExcelSolution
├── ExcelFiles
│   ├── Example.xlsx               # Sample file for data extraction
│   ├── ExampleForDataExport.xlsx   # File generated during export
├── Objects
│   └── Users.cs                   # User data model representing a user in the system
├── Services
│   └── ExcelManager.cs             # Excel helper functions (file checking, path management, worksheet validation)
├── Utils
│   ├── ExportNewExcel.cs           # Logic for exporting manipulated user data to Excel
│   └── ExtractDataFromExcel.cs     # Logic for extracting user data from Excel
├── Program.cs                      # Main program entry point, orchestrates extraction and export
```

### Key Features:
- Extract data from an existing Excel file (`Example.xlsx`).
- Manipulate the extracted data, such as generating random values for certain fields.
- Export the manipulated data to a new Excel file (`ExampleForDataExport.xlsx`).
- Easy-to-follow code structure with helper functions for file management.



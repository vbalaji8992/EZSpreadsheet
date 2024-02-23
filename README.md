# EZSpreadsheet
EZSpreadsheet is an easy-to-use .NET library to work with spreadsheet files. It is built as a wrapper around the more complex **OpenXML** library.

## Installation
To include this library in your project, search for `EZSpreadsheet` in Visual Studio NuGet Package Manager window.

**NuGet link**
https://www.nuget.org/packages/EZSpreadsheet/

## Usage
Sample C# code

    // Create a workbook in the given path
    EZWorkbook workbook = new("Output/Example.xlsx");

    // Add a worksheet
    EZWorksheet worksheet = workbook.AddSheet("Sheet1");

    // Insert text in cell A1
    worksheet.GetCell(1, 1).SetValue("EZSpreadsheet Example");

    // Save workbook
    workbook.Save();
    
Refer file `Program.cs` in the folder `EZSpreadsheet.Examples` for more examples.


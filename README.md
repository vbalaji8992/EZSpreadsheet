# EZSpreadsheet
EZSpreadsheet is an easy-to-use .NET library to work with spreadsheet files. It is built as a wrapper around the more complex OpenXML library.

## Installation
To include this library in your project, search for `EZSpreadsheet` in Visual Studio NuGet Package Manager window or run the following command in the Package Manager Console

    PM> Install-Package EZSpreadsheet 

## Usage
The following code creates a workbook named `Example` with a worksheet named `Sheet1` and inserts the text `EZSpreadsheet Example` in the cell `A1`

    EZWorkbook workbook = new("Example.xlsx");
    EZWorksheet worksheet = workbook.AddSheet("Sheet1");
    worksheet.GetCell(1, 1).SetValue("EZSpreadsheet Example");
    workbook.Save();
    
Refer file `Program.cs` in the folder `EZSpreadsheet.Examples` for more examples.

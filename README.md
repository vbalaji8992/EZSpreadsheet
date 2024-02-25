# EZSpreadsheet
EZSpreadsheet is an easy-to-use .NET library to work with spreadsheet files. It is built as a wrapper around the more complex **OpenXML** library.

## Installation
To include this library in your project, search for `EZSpreadsheet` in Visual Studio NuGet Package Manager window.

**NuGet link**
https://www.nuget.org/packages/EZSpreadsheet/

## Features

 - [x] Create spreadsheet
 - [x] Add text, number and formula
 - [x] Format text and number
 - [x] Insert list of primitives
 - [x] Insert list of objects
 - [x] Apply styling to cells
 - [ ] Align content
 - [ ] Merge cells
 - [ ] Set row height and column width
 - [ ] Create charts
 - [ ] Open and edit spreadsheet

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
    
More examples here
[Wiki](https://github.com/vbalaji8992/EZSpreadsheet/wiki)
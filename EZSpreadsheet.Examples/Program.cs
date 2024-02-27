// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;
using EZSpreadsheet.Style;

// Create new workbook in the path
EZWorkbook workbook = new("Output/EzBook.xlsx");

// Create new worksheet with the given name
EZWorksheet worksheet = workbook.AddSheet("EzSheet");

// Set content of cell A1 as string
worksheet.GetCell(1, 1).SetValue("Decimals");

// Set content of cell A2 as integer
worksheet.GetCell("A2").SetValue(123);

// Set content of cell A3 as decimal
worksheet.GetCell("A", 3).SetValue(12.34);

// Set content of cell A4 as decimal and as per given format Id
// Refer the below link to look-up format Ids
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat
worksheet.GetCell("A4")
    .SetValue(12.3456)
    .SetStyle(new EZStyle { NumberFormatId = 2 });

// Set formula of cell A5 
worksheet.GetCell("A5").SetFormula("SUM(A2:A4)");

// Set content of cell A7 as string and with the given fill and font color
worksheet.GetCell("A7")
    .SetValue("Integers")
    .SetStyle(new EZStyle { FillColor = "#FFFF00", FontColor = "#FF0000" });

// Insert list of integers starting from cell A8
var list1 = new List<int> { 1, 2, 3 };
var range = worksheet.GetCell("a8").InsertData(list1);
// The method InsertData will return the cell range in which the list was inserted
// Set the font size of the above range
range.SetStyle(new EZStyle { FontSize = 14 });

// Insert list of type Tuple at cell C2 and apply border and font
var list2 = new List<Tuple<string, double, int>>()
{
    Tuple.Create("Jack", 78.8, 8),
    Tuple.Create("Abbey", 92.1, 9),
    Tuple.Create("Dave", 88.3, 9),
    Tuple.Create("Sam", 91.7, 8),
    Tuple.Create("Ed", 71.2, 5)
};
// Create a style object with custom border and font
var tableStyle = new EZStyle { BorderType = EZBorder.Thin, Font = EZFont.TimesNewRoman };
// Insert the above list with property name as heading and apply the above style
worksheet.GetCell("C2")
    .InsertData(list2, new EZListOptions { AddPropertyNameAsHeading = true })
    .SetStyle(tableStyle);

// Set the contents of cell C9 as string and make it bold
worksheet.GetCell("C9")
    .SetValue("Transposed Integers")
    .SetStyle(new EZStyle { IsBold = true });

// Insert list of integers from left to right starting from cell C10 
worksheet.GetCell("C10")
    .InsertData(list1, new EZListOptions { TransposeData = true });

// Save the workbook
workbook.Save();
// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;
using EZSpreadsheet.Style;

// Create new workbook in the path
EZWorkbook workbook = new("Output/EzBook.xlsx");

// Create new worksheet
EZWorksheet worksheet = workbook.AddSheet("EzSheet");

// Set content of cell A1 as string
worksheet.GetCell(1, 1).SetValue("Heading");

// Set content of cell A2 as integer
worksheet.GetCell("A2").SetValue(123);

// Set content of cell A3 as decimal/double
worksheet.GetCell("A", 3).SetValue(12.34);

// Set content of cell A4 as decimal/double and as per given format Id (Refer the below link to look-up format Ids)
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat
worksheet.GetCell("A4")
    .SetValue(12.3456)
    .SetStyle(new EZStyle { NumberFormatId = 2 });

// Set formula of cell A5 
worksheet.GetCell("A5").SetFormula("SUM(A2:A4)");

// Set content of cell A7 as string and with the given fill and font color
worksheet.GetCell("A7")
    .SetValue("Numbers")
    .SetStyle(new EZStyle { FillColor = EZColor.Yellow, FontColor = EZColor.Red });

// Insert list of integers starting from cell A8
// The method will return the cell range in which the list was inserted
var list1 = new List<int> { 1, 2, 3 };
var range = worksheet.GetCell("a8").InsertData(list1);
// Fill the above range with the given color
range.SetStyle(new EZStyle { FillColor = EZColor.Pink });


// Create a list of integers of type Data
var list2 = new List<Data>();
for (uint i = 1; i <= 5; i++)
{
    list2.Add(new Data()
    {
        Prop1 = new Random().Next(100),
        Prop2 = new Random().Next(100),
        Prop3 = new Random().Next(100)
    });
}
// Create a style object with custom border and font
var tableStyle = new EZStyle { BorderType = EZBorder.Thin, Font = EZFont.TimesNewRoman };
// Insert the above list with property name as heading and apply the above style
range = worksheet.GetCell("A12")
    .InsertData(list2, true)
    .SetStyle(tableStyle);

// Save the workbook
workbook.Save();
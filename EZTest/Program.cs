// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;

EZWorkbook workbook = new("EzBook.xlsx");
var worksheet = workbook.AddSheet("EZ");

for (uint i = 1; i < 10000; i++)
{
    worksheet.GetCell(i, 1).SetValue((int)i);
}

worksheet.GetCell(1,1).ApplyStyle(1);

workbook.Save();
// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;

EZWorkbook workbook = new("EzBook.xlsx");
var worksheet = workbook.AddSheet("EZ");

List<uint> list = new List<uint>();

for (uint i = 1; i < 10000; i++)
{
    list.Add(i);
}

worksheet.InsertData(list, "C4");

workbook.Save();
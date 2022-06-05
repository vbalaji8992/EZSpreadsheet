// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;

EZWorkbook workbook = new("EzBook.xlsx");
var worksheet = workbook.AddSheet("EZ");

List<uint> list = new List<uint>();

for (uint i = 1; i < 100000; i++)
{
    list.Add(i);
}

worksheet.InsertData(list, "C4");
//worksheet.GetCell(4, 1).SetText("a");
//worksheet.GetCell(4, 3).ConvertToNumber();
//worksheet.GetCell(1, 1).SetValue(1);

workbook.Save();
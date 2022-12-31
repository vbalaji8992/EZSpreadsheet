// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;

EZWorkbook workbook = new("EzBook.xlsx");
var worksheet = workbook.AddSheet("EZ");

List<Data> list = new List<Data>();

for (uint i = 1; i < 10; i++)
{
    list.Add(new Data()
    {
        Prop1 = new Random().Next(100),
        Prop2 = new Random().Next(100),
        Prop3 = new Random().Next(100)
    });
}
var range1 = worksheet.GetCell("a4").InsertData(list, true);
range1?.ConvertToNumber();

var range2 = worksheet.GetCell("a15").InsertData(new List<int>(), true);

var range3 = worksheet.GetCell("a16").InsertData(new List<int>() { 1, 2, 3}, true);
range3?.SetFontStyle(new EZFontStyle() { IsBold = true });

worksheet.GetCell("a", 1).SetText("a");
//worksheet.GetCell(4, 3).ConvertToNumber();
//worksheet.GetCell(1, 1).SetValue(1);

workbook.Save();

class Data
{
    public int Prop1 { get; set; }
    public int Prop2 { get; set; }
    public int Prop3 { get; set; }
}
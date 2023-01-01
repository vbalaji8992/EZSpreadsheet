// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;
using EZSpreadsheet.StyleEnums;

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
var range1 = worksheet.GetCell("b2").InsertData(list, true);
range1?.ConvertToNumber();
range1?.SetStyle(new EZCellStyle { BorderType = EZBorder.Thin, FillColor = EZColor.Yellow });

worksheet.GetCell("a", 1).SetText("a1");
worksheet.GetCell("a", 1).SetStyle(new EZCellStyle { FillColor = EZColor.Green });

workbook.Save();

Console.WriteLine("Done");
Console.ReadLine();

class Data
{
    public int Prop1 { get; set; }
    public int Prop2 { get; set; }
    public int Prop3 { get; set; }
}
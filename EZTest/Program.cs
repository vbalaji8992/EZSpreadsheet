// See https://aka.ms/new-console-template for more information
using EZSpreadsheet;
using EZSpreadsheet.Style;

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
worksheet.GetCell("b2")
    .InsertData(list, true)
    .SetStyle(new EZStyle { BorderType = EZBorder.Thin, FillColor = EZColor.Yellow, Font = EZFont.TimesNewRoman });

worksheet.GetCell("a", 1)
    .SetValue<double?>(12.345678)
    .SetStyle(new EZStyle { FillColor = EZColor.Green, NumberFormatId = 2, Font = EZFont.Arial });

worksheet.GetCell("a", 2)
    .SetValue<string?>(null)
    .SetStyle(new EZStyle { BorderType = EZBorder.Thin, FillColor = EZColor.Yellow, Font = EZFont.Century, FontSize = 50 });

worksheet.GetCell("f6")
    .SetFormula("SUM(B3:D11)");

worksheet.GetCell("f2")
    .SetFormula("CONCATENATE(B2,C2,D2)");

worksheet.GetRange("j11", "e2").SetStyle(new EZStyle { FillColor = EZColor.Pink });

workbook.Save();

Console.WriteLine("Done");
Console.ReadLine();

class Data
{
    public int? Prop1 { get; set; }
    public int Prop2 { get; set; }
    public int Prop3 { get; set; }
}
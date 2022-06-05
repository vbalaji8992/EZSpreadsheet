using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EZSpreadsheet.StyleEnums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EZSpreadsheet
{
    internal class EZStylesheet
    {
        EZWorkbook WorkBook { get; }
        public WorkbookStylesPart WorkbookStylesPart { get; }

        public EZStylesheet(EZWorkbook workBook, WorkbookStylesPart workbookStylesPart)
        {
            WorkBook = workBook;
            WorkbookStylesPart = workbookStylesPart;
            WorkbookStylesPart.Stylesheet = new Stylesheet();
            AppendBasicStyles();
        }

        void AppendBasicStyles()
        {
            var fonts = new Fonts();
            fonts.Append(new Font()
            {
                FontSize = new FontSize() { Val = 11 },
                Color = new Color() { Indexed = (uint)EZFontColor.Black },
                FontName = new FontName() { Val = EZFont.Calibri.ToString() }
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            var fills = new Fills();
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            });
            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            borders.Append(new Border());

            var cellFormats = new CellFormats();
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0,
                FormatId = 0
            });           
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            WorkbookStylesPart.Stylesheet.Append(fonts);
            WorkbookStylesPart.Stylesheet.Append(fills);
            WorkbookStylesPart.Stylesheet.Append(borders);
            WorkbookStylesPart.Stylesheet.Append(cellFormats);
        }
    }
}

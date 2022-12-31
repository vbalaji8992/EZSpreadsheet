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
        WorkbookStylesPart WorkbookStylesPart { get; }
        Dictionary<EZFontStyle, uint> FontIndices { get; }
        Dictionary<uint, uint> StyleIndices { get; }

        private Fonts fonts;
        private Fills fills;
        private Borders borders;
        private CellFormats cellFormats;

        public EZStylesheet(EZWorkbook workBook, WorkbookStylesPart workbookStylesPart)
        {
            WorkBook = workBook;
            WorkbookStylesPart = workbookStylesPart;
            WorkbookStylesPart.Stylesheet = new Stylesheet();
            FontIndices = new Dictionary<EZFontStyle, uint>();
            StyleIndices = new Dictionary<uint, uint>();
            fonts = new Fonts();
            fills = new Fills();
            borders = new Borders();
            cellFormats = new CellFormats();
            AppendBasicStyles();
        }

        void AppendBasicStyles()
        {           
            var fontId = AppendFontStyle(new EZFontStyle());

            fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            });
            fills.Count = (uint)fills.ChildElements.Count;

            borders.Append(new Border());

            AppendCellFormat(fontId);

            WorkbookStylesPart.Stylesheet.Append(fonts);
            WorkbookStylesPart.Stylesheet.Append(fills);
            WorkbookStylesPart.Stylesheet.Append(borders);
            WorkbookStylesPart.Stylesheet.Append(cellFormats);
        }

        public uint AppendFontStyle(EZFontStyle fontStyle)
        {
            var existingStyle = FontIndices.FirstOrDefault(kvp => kvp.Key == fontStyle);
            if (existingStyle.Key != null)
                return existingStyle.Value;

            fonts.Append(new Font()
            {
                FontSize = new FontSize() { Val = fontStyle.FontSize },
                Color = new Color() { Indexed = (uint)fontStyle.FontColor },
                FontName = new FontName() { Val = fontStyle.Font.ToString() },
                Bold = (fontStyle.IsBold)? new Bold() : null,
                Italic = (fontStyle.IsItalic)? new Italic() : null,
                Underline = (fontStyle.IsUnderlined)? new Underline(): null
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            uint fontId = fonts.Count - 1;
            FontIndices.Add(fontStyle, fontId);
            return fontId;
        }

        public uint AppendCellFormat(uint fontId)
        {
            if (StyleIndices.ContainsKey(fontId))
                return StyleIndices[fontId];

            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = fontId,
                NumberFormatId = 0,
                FormatId = 0,
                ApplyFont = true
            });
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            uint styleIndex = cellFormats.Count - 1;
            StyleIndices.Add(fontId, styleIndex);
            return styleIndex;
        }
    }
}

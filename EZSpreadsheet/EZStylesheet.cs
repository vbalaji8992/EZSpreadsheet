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
        List<EZCellStyle> CellStyleList { get; }
        Dictionary<EZCellStyle, uint> CellStyleIndex { get; }

        private Fonts fonts;
        private Fills fills;
        private Borders borders;
        private CellFormats cellFormats;

        public EZStylesheet(EZWorkbook workBook, WorkbookStylesPart workbookStylesPart)
        {
            WorkBook = workBook;
            WorkbookStylesPart = workbookStylesPart;
            WorkbookStylesPart.Stylesheet = new Stylesheet();
            CellStyleList = new List<EZCellStyle>();
            CellStyleIndex = new Dictionary<EZCellStyle, uint>();
            fonts = new Fonts();
            fills = new Fills();
            borders = new Borders();
            cellFormats = new CellFormats();
            AppendBasicStyles();
        }

        void AppendBasicStyles()
        {           
            var cellStyle = AppendCellStyle(new EZCellStyle());

            AppendCellFormat(cellStyle);

            WorkbookStylesPart.Stylesheet.Append(fonts);
            WorkbookStylesPart.Stylesheet.Append(fills);
            WorkbookStylesPart.Stylesheet.Append(borders);
            WorkbookStylesPart.Stylesheet.Append(cellFormats);
        }

        public EZCellStyle AppendCellStyle(EZCellStyle cellStyle)
        {
            var existingStyle = CellStyleList.FirstOrDefault(x => cellStyle.Equals(x));
            if (existingStyle != null)
                return existingStyle;

            var fontMatch = CellStyleList.FirstOrDefault(x => cellStyle.FontEquals(x));
            if (fontMatch == null)
                cellStyle.FontId = AppendFont(cellStyle);
            else
                cellStyle.FontId = fontMatch.FontId;

            var borderMatch = CellStyleList.FirstOrDefault(x => cellStyle.BorderEquals(x));
            if (borderMatch == null)
                cellStyle.BorderId = AppendBorder(cellStyle);
            else
                cellStyle.BorderId = borderMatch.BorderId;

            var fillMatch = CellStyleList.FirstOrDefault(x => cellStyle.FillEquals(x));
            if (fillMatch == null)
                cellStyle.FillId = AppendFill(cellStyle);
            else
                cellStyle.FillId = fillMatch.FillId;

            CellStyleList.Add(cellStyle);
            return cellStyle;
        }

        private uint AppendFont(EZCellStyle cellStyle)
        {
            fonts.Append(new Font()
            {
                FontSize = new FontSize() { Val = cellStyle.FontSize },
                Color = new Color() { Indexed = (uint)cellStyle.FontColor },
                FontName = new FontName() { Val = cellStyle.Font.ToString() },
                Bold = (cellStyle.IsBold) ? new Bold() : null,
                Italic = (cellStyle.IsItalic) ? new Italic() : null,
                Underline = (cellStyle.IsUnderlined) ? new Underline() : null
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            return fonts.Count - 1;
        }

        private uint AppendBorder(EZCellStyle cellStyle)
        {
            var border = new Border();

            LeftBorder leftBorder = new LeftBorder() { Style = (BorderStyleValues)cellStyle.BorderType };
            RightBorder rightBorder = new RightBorder() { Style = (BorderStyleValues)cellStyle.BorderType };
            TopBorder topBorder = new TopBorder() { Style = (BorderStyleValues)cellStyle.BorderType };            
            BottomBorder bottomBorder = new BottomBorder() { Style = (BorderStyleValues)cellStyle.BorderType };

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);            
            border.Append(bottomBorder);

            borders.Append(border);

            borders.Count = (uint)borders.ChildElements.Count;
            return borders.Count - 1;
        }

        private uint AppendFill(EZCellStyle cellStyle)
        {
            if (fills.ChildElements.Count == 0)
            {
                fills.Append(new Fill()
                {
                    PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 }
                });
                fills.Append(new Fill()
                {
                    PatternFill = new PatternFill() { PatternType = PatternValues.None }
                });
            }
            else
            {
                fills.Append(new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor() { Indexed = (uint)cellStyle.FillColor }
                    }
                });                
            }

            fills.Count = (uint)fills.ChildElements.Count;
            return fills.Count - 1;
        }

        public uint AppendCellFormat(EZCellStyle cellStyle)
        {
            var existingStyle = CellStyleIndex.FirstOrDefault(kvp => cellStyle.Equals(kvp.Key));
            if (existingStyle.Key != null)
                return existingStyle.Value;

            cellFormats.Append(new CellFormat()
            {
                BorderId = cellStyle.BorderId,
                FillId = cellStyle.FillId,
                FontId = cellStyle.FontId,
                NumberFormatId = cellStyle.NumberFormatId,
                FormatId = 0,
                ApplyFont = true,
                ApplyBorder = true,
                ApplyFill = true
            });
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            uint styleIndex = cellFormats.Count - 1;
            CellStyleIndex.Add(cellStyle, styleIndex);
            return styleIndex;
        }
    }
}

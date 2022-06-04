﻿using DocumentFormat.OpenXml.Packaging;
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
                FontSize = new FontSize() { Val = 16 },
                Color = new Color() { Indexed = (uint)EZFontColor.Black },
                FontName = new FontName() { Val = EZFont.Calibri.ToString() }
                //FontFamilyNumbering = new FontFamilyNumbering() { Val = 0 },
                //FontScheme = new FontScheme() { Val = FontSchemeValues.Minor }
            });
            fonts.Count = (uint)fonts.ChildElements.Count;

            var fills = new Fills();
            fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            });
            //fills.Append(new Fill()
            //{
            //    PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 }
            //});
            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            borders.Append(new Border());

            //var cellStyleFormats = new CellStyleFormats();
            //cellStyleFormats.Append(new CellFormat()
            //{
            //    NumberFormatId = 0,
            //    FontId = 0,
            //    FillId = 0,
            //    BorderId = 0
            //});
            //cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;

            var cellFormats = new CellFormats();
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0,
                FormatId = 0
            });
            cellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 2,
                FormatId = 0,
                ApplyNumberFormat = true
            });
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            //var cellStyles = new CellStyles();
            //cellStyles.Append(new CellStyle()
            //{
            //    Name = "Normal",
            //    FormatId = 0,
            //    BuiltinId = 0
            //});
            //cellStyles.Count = (uint)cellStyles.ChildElements.Count;

            WorkbookStylesPart.Stylesheet.Append(fonts);
            WorkbookStylesPart.Stylesheet.Append(fills);
            WorkbookStylesPart.Stylesheet.Append(borders);
            //WorkbookStylesPart.Stylesheet.Append(cellStyleFormats);
            WorkbookStylesPart.Stylesheet.Append(cellFormats);
            //WorkbookStylesPart.Stylesheet.Append(cellStyles);
            //WorkbookStylesPart.Stylesheet.Save();
        }

        internal void GetByStyleIndex(uint index)
        {
            var cellFormat = WorkbookStylesPart.Stylesheet.CellFormats?.Elements<CellFormat>().Skip((int)index).First();
            Console.WriteLine(cellFormat.NumberFormatId.ToString());
        }
    }
}

using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EZSpreadsheet.Style;

namespace EZSpreadsheet.Tests
{
    public class CellStyleTest
    {
        public CellStyleTest()
        {

        }

        [Fact]
        public void ShouldSetFontToCell()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { Font = EZFont.Arial });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { Font = EZFont.Calibri });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { Font = EZFont.Century });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetFontToCell.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldSetFontColor()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { FontColor = EZColor.Red });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { FontColor = EZColor.Green });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { FontColor = EZColor.Blue });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetFontColor.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldSetFontSize()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { FontSize = 8 });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { FontSize = 12 });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { FontSize = 16 });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetFontSize.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldMakeTextBold()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { IsBold = true });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { IsBold = true });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { IsBold = true });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldMakeTextBold.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldMakeTextItalic()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { IsItalic = true });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { IsItalic = true });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { IsItalic = true });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldMakeTextItalic.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldMakeTextUnderlined()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { IsUnderlined = true });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { IsUnderlined = true });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { IsUnderlined = true });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldMakeTextUnderlined.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShoulSetCellBorderWithColor()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetStyle(new EZStyle { BorderType = EZBorder.Thin, BorderColor = EZColor.Black });
            ws.GetCell(3, 1).SetStyle(new EZStyle { BorderType = EZBorder.Thick, BorderColor = EZColor.Red });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShoulSetCellBorderWithColor.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShoulFillCellWithColor()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetStyle(new EZStyle { FillColor = EZColor.Yellow });
            ws.GetCell(3, 1).SetStyle(new EZStyle { FillColor = EZColor.Pink });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShoulFillCellWithColor.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldFormatNumbers()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue(123.45).SetStyle(new EZStyle { NumberFormatId = 1 });
            ws.GetCell(2, 1).SetValue(123.456).SetStyle(new EZStyle { NumberFormatId = 2 });
            ws.GetCell(3, 1).SetValue(1.23).SetStyle(new EZStyle { NumberFormatId = 10 });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldFormatNumbers.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldApplyMultipleStylesToCell()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            EZStyle cellStyle = new EZStyle 
            {
                Font = EZFont.Arial,
                FontColor = EZColor.Red,
                FontSize = 8,
                IsBold = true,
                BorderType = EZBorder.Thin,
                BorderColor = EZColor.Black,
                FillColor = EZColor.Yellow,
                NumberFormatId = 1 
            };
            ws.GetCell(2, 2).SetValue(123).SetStyle(cellStyle);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldApplyMultipleStylesToCell.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldApplyStylesToRangeOfCells()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            EZStyle cellStyle = new EZStyle
            {
                BorderType = EZBorder.Thin,
                BorderColor = EZColor.Black,
                FillColor = EZColor.Yellow
            };
            ws.GetRange("b2", "f5").SetStyle(cellStyle);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldApplyStylesToRangeOfCells.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }
    }
}

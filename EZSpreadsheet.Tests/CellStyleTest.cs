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
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new EZStyle { FontColor = "#FF0800" });
            ws.GetCell(2, 1).SetValue(12345).SetStyle(new EZStyle { FontColor = "#17FF00" });
            ws.GetCell(3, 1).SetValue(123.45).SetStyle(new EZStyle { FontColor = "#0042FF" });
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
            ws.GetCell(1, 1).SetStyle(new EZStyle { BorderType = EZBorder.Thin, BorderColor = "#000000" });
            ws.GetCell(3, 1).SetStyle(new EZStyle { BorderType = EZBorder.Thick, BorderColor = "#FF0800" });
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
            ws.GetCell(1, 1).SetStyle(new EZStyle { FillColor = "#ECFF00" });
            ws.GetCell(3, 1).SetStyle(new EZStyle { FillColor = "#FF00EC" });
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
                FontColor = "#FF0800",
                FontSize = 8,
                IsBold = true,
                BorderType = EZBorder.Thin,
                BorderColor = "#000000",
                FillColor = "#ECFF00",
                NumberFormatId = 1 
            };
            ws.GetCell(2, 2).SetValue(123).SetStyle(cellStyle);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldApplyMultipleStylesToCell.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ColorCodesShouldBeCaseInsensitive()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            EZStyle cellStyle = new EZStyle
            {
                Font = EZFont.Arial,
                FontColor = "#ff0800",
                FontSize = 8,
                IsBold = true,
                BorderType = EZBorder.Thin,
                BorderColor = "#000000",
                FillColor = "#ecff00",
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
                BorderColor = "#000000",
                FillColor = "#ECFF00"
            };
            ws.GetRange("b2", "f5").SetStyle(cellStyle);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldApplyStylesToRangeOfCells.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldNotApplyColorForInvalidColorCodes()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            EZStyle cellStyle = new EZStyle
            {
                BorderType = EZBorder.Thin,
                BorderColor = "#123ABC123",
                FillColor = "#ABC123ABC"
            };
            ws.GetRange("b2", "f5").SetStyle(cellStyle);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldNotApplyColorForInvalidColorCode.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }
    }
}

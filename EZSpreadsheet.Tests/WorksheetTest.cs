using Xunit.Abstractions;
using Moq;
using EZSpreadsheet;

namespace EZSpreadsheet.Tests
{
    public class WorksheetTest
    {
        public WorksheetTest()
        {
        }

        [Fact]
        public void ShouldThrowExceptionWhenGettingCellForInvalidRowIndex()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            var worksheet = workbook.AddSheet("sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A", 0));
        }

        [Fact]
        public void ShouldThrowExceptionWhenGettingCellForInvalidColumnName()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            var worksheet = workbook.AddSheet("sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("AAAAA", 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell(1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell(1, 100000));
        }        

        [Fact]
        public void ShouldThrowExceptionForInvalidCellReference()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            var worksheet = workbook.AddSheet("sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell(""));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("1"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("1A"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("AAAA1"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A-1"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A0"));
        }

        [Fact]
        public void ShouldAddCellsInSameColumn1() 
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");           

            ws.GetCell(1, 1);
            ws.GetCell("a", 2);
            ws.GetCell("A", 2);
            ws.GetCell("a4");
            ws.GetCell("A5");

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddCellsInSameColumn.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldAddCellsInSameColumn2()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            for (int i = 0; i < 10; i++)
            {
                ws.GetCell(1, 1);
                ws.GetCell("a", 2);
                ws.GetCell("A", 2);
                ws.GetCell("a4");
                ws.GetCell("A5"); 
            }

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddCellsInSameColumn.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldAddCellsInSameRow()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            ws.GetCell(1, 1);
            ws.GetCell("b", 1);
            ws.GetCell("C", 1);
            ws.GetCell("d1");
            ws.GetCell("E1");

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddCellsInSameRow.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldRandomlyAddCells()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            ws.GetCell(10, 3);
            ws.GetCell("b", 1);
            ws.GetCell("C", 1);
            ws.GetCell("f6");
            ws.GetCell("A1");

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldRandomlyAddCells.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldAddCellsinColumnsAfterZ()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            ws.GetCell(1, 1);
            ws.GetCell(1, 26);
            ws.GetCell(1, 100);

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddCellsinColumnsAfterZ.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldAddRangeOfCells1()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws1 = wb.AddSheet("sheet1");
            var ws2 = wb.AddSheet("sheet2");

            ws1.GetRange("a1", "j10");
            ws2.GetRange("j10", "a1");

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddRangeOfCells.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldAddRangeOfCells2()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws1 = wb.AddSheet("sheet1");
            var ws2 = wb.AddSheet("sheet2");

            ws1.GetRange(1, 1, 10, 10);
            ws2.GetRange(10, 10, 1, 1);

            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddRangeOfCells.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldGetFirstAndLastRowIndex()
        {
            var wb = new EZWorkbook(new MemoryStream());
            var ws = wb.AddSheet("sheet1");

            ws.GetCell("D4");
            ws.GetCell("A34");

            Assert.Equal((uint)4, ws.GetFirstRowIndex());
            Assert.Equal((uint)34, ws.GetLastRowIndex());
        }

        [Fact]
        public void ShouldGetFirstAndLastColumnIndex()
        {
            var wb = new EZWorkbook(new MemoryStream());
            var ws = wb.AddSheet("sheet1");

            ws.GetCell("J4");
            ws.GetCell("C34");

            Assert.Equal((uint)3, ws.GetFirstColumnIndex());
            Assert.Equal((uint)10, ws.GetLastColumnIndex());
        }

        [Fact]
        public void ShouldGetFirstAndLastColumnName()
        {
            var wb = new EZWorkbook(new MemoryStream());
            var ws = wb.AddSheet("sheet1");

            ws.GetCell("J4");
            ws.GetCell("C34");

            Assert.Equal("C", ws.GetFirstColumnName());
            Assert.Equal("J", ws.GetLastColumnName());
        }

        [Fact]
        public void ShouldThrowExceptionIfSheetIsEmpty()
        {
            var wb = new EZWorkbook(new MemoryStream());
            var ws = wb.AddSheet("sheet1");

            Assert.Throws<Exception>(() => ws.GetFirstRowIndex());
            Assert.Throws<Exception>(() => ws.GetLastRowIndex());
            Assert.Throws<Exception>(() => ws.GetFirstColumnIndex());
            Assert.Throws<Exception>(() => ws.GetLastColumnIndex());
            Assert.Throws<Exception>(() => ws.GetFirstColumnName());
            Assert.Throws<Exception>(() => ws.GetLastColumnName());
        }

        [Fact]
        public void ShouldGetSheetName()
        {
            var wb = new EZWorkbook(new MemoryStream());
            var ws = wb.AddSheet("NewSheet1234");

            Assert.Equal("NewSheet1234", ws.GetSheetName());
        }
    }
}

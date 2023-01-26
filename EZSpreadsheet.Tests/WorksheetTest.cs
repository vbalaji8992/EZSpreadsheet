using Xunit.Abstractions;
using Moq;
using EZSpreadsheet;

namespace EZSpreadsheet.Tests
{
    public class WorksheetTest
    {
        public WorksheetTest()
        {
            TestHelper.CreateFolder(TestHelper.TEST_OUTPUT_FOLDER);
        }

        [Fact]
        public void ShouldThrowExceptionWhenGettingCellForInvalidRowIndex()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldThrowExceptionWhenGettingCellForInvalidRowIndex.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A", 0));
        }

        [Fact]
        public void ShouldThrowExceptionWhenGettingCellForInvalidColumnName()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldThrowExceptionWhenGettingCellForInvalidColumnName.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("AAAAA", 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell(1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell(1, 100000));
        }

        [Fact]
        public void ShouldAddCellIfNotExists()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldAddCellIfNotExists.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            var cell = worksheet.GetCell("A", 1);

            Assert.NotNull(cell);
        }

        [Fact]
        public void ShouldReturnCellIfAlreadyExists()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldReturnCellIfAlreadyExists.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            worksheet.GetCell("A", 1);
            var cell = worksheet.GetCell("A", 1);

            Assert.NotNull(cell);
        }

        [Fact]
        public void ShouldGetCellForGivenRowAndColumnIndex()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldGetCellForGivenRowAndColumnIndex.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            var cell = worksheet.GetCell(1, 1);

            Assert.NotNull(cell);
        }

        [Fact]
        public void ShouldGetCellForGivenCellReference()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldGetCellForGivenCellReference.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            var cell = worksheet.GetCell("A1");

            Assert.NotNull(cell);
        }

        [Fact]
        public void ShouldThrowExceptionForInvalidCellReference()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldThrowExceptionForInvalidCellReference.xlsx");
            var worksheet = workbook.AddSheet("sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell(""));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("1"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("1A"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("AAAA1"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A-1"));
            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A0"));
        }
    }
}

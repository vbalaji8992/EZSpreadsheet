using Xunit.Abstractions;
using Moq;
using EZSpreadsheet;

namespace EZSpreadsheet.Tests
{
    public class WorksheetTest
    {
        private const string TEST_RESOURCES_FOLDER = "TestResources/WorksheetTest";

        public WorksheetTest()
        {
            if (!Directory.Exists(TEST_RESOURCES_FOLDER))
            {
                Directory.CreateDirectory(TEST_RESOURCES_FOLDER);
            }            
        }

        [Fact]
        public void ShouldThrowExceptionWhenGettingCellForInvalidRowIndex()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldThrowExceptionWhenGettingCellForInvalidRowIndex.xlsx");
            var worksheet = new EZWorksheet(workbook, "sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("A", 0));
        }

        [Fact]
        public void ShouldThrowExceptionWhenGettingCellForInvalidColumnName()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldThrowExceptionWhenGettingCellForInvalidColumnName.xlsx");
            var worksheet = new EZWorksheet(workbook, "sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => worksheet.GetCell("AAAAA", 1));
        }

        [Fact]
        public void ShouldAddCellIfNotExists()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldAddCellIfNotExists.xlsx");
            var worksheet = new EZWorksheet(workbook, "sheet1");

            var cell = worksheet.GetCell("A", 1);

            Assert.NotNull(cell);
        }

        [Fact]
        public void ShouldReturnCellIfAlreadyExists()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldReturnCellIfAlreadyExists.xlsx");
            var worksheet = new EZWorksheet(workbook, "sheet1");

            worksheet.GetCell("A", 1);
            var cell = worksheet.GetCell("A", 1);

            Assert.NotNull(cell);
        }
    }
}

using Xunit.Abstractions;

namespace EZSpreadsheet.Tests
{
    public class WorkbookTest
    {
        private readonly ITestOutputHelper output;
        private const string TEST_RESOURCES_FOLDER = "TestResources/WorkbookTest";

        public WorkbookTest(ITestOutputHelper output)
        {
            this.output = output;

            if (!Directory.Exists(TEST_RESOURCES_FOLDER))
            {
                Directory.CreateDirectory(TEST_RESOURCES_FOLDER);
            }
        }

        [Fact]
        public void ShouldCreateWorkbookWithNoSheets()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldCreateWorkbookWithNoSheets.xlsx");            
            Assert.Equal(0, workbook.GetSheetCount());
        }

        [Fact]
        public void ShouldAddSheetWithDefaultName()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldAddSheetWithDefaultName.xlsx");
            workbook.AddSheet();
            Assert.NotNull(workbook.GetSheet("Sheet1"));
        }

        [Fact]
        public void ShouldAddSheetWithGivenName()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldAddSheetWithGivenName.xlsx");
            workbook.AddSheet("NewSheet");
            Assert.NotNull(workbook.GetSheet("NewSheet"));
        }

        [Fact]
        public void ShouldThrowExceptionWhenAddingSheetIfNameAlreadyExists()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldThrowExceptionWhenAddingSheetIfNameAlreadyExists.xlsx");
            workbook.AddSheet("NewSheet");
            Assert.Throws<Exception>(() => workbook.AddSheet("NewSheet"));
        }

        [Fact]
        public void ShouldGetSheetByNameIfExists()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldGetSheetByNameIfExists.xlsx");
            workbook.AddSheet("NewSheet");
            var sheet = workbook.GetSheet("NewSheet");
            Assert.Equal("NewSheet", sheet?.Sheet.Name);
        }

        [Fact]
        public void ShouldReturnNullIfSheetDoesNotExist()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldReturnNullIfSheetDoesNotExist.xlsx");
            workbook.AddSheet("NewSheet");
            Assert.Null(workbook.GetSheet("OldSheet"));
        }

        [Fact]
        public void ShouldGetSheetCount()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldGetSheetCount.xlsx");
            workbook.AddSheet("Sheet1");
            workbook.AddSheet("Sheet2");
            workbook.AddSheet("Sheet3");
            Assert.Equal(3, workbook.GetSheetCount());
        }

        [Fact]
        public void ShouldSaveAndCloseWorkbook()
        {
            var workbook = new EZWorkbook($"{TEST_RESOURCES_FOLDER}/ShouldSaveAndCloseWorkbook.xlsx");
            workbook.AddSheet("Sheet1");
            workbook.Save();
            File.Delete($"{TEST_RESOURCES_FOLDER}/ShouldSaveAndCloseWorkbook.xlsx");
        }
    }
}
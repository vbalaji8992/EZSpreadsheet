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

        [Fact]
        public void ShouldAddCellsInSameColumn1() 
        {
            var testName = "ShouldAddCellsInSameColumn";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");           

            ws.GetCell(1, 1);
            ws.GetCell("a", 2);
            ws.GetCell("A", 2);
            ws.GetCell("a4");
            ws.GetCell("A5");

            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}\{testName}.xml";

            TestHelper.AssertFile(expectedXmlFile, $@"{extractPath}\xl\worksheets\sheet1.xml");
        }

        [Fact]
        public void ShouldAddCellsInSameColumn2()
        {
            var testName = "ShouldAddCellsInSameColumn";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}.xlsx";
            var wb = new EZWorkbook(file);
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

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}\{testName}.xml";

            TestHelper.AssertFile(expectedXmlFile, $@"{extractPath}\xl\worksheets\sheet1.xml");
        }

        [Fact]
        public void ShouldAddCellsInSameRow()
        {
            var testName = "ShouldAddCellsInSameRow";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");

            ws.GetCell(1, 1);
            ws.GetCell("b", 1);
            ws.GetCell("C", 1);
            ws.GetCell("d1");
            ws.GetCell("E1");

            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}\{testName}.xml";

            TestHelper.AssertFile(expectedXmlFile, $@"{extractPath}\xl\worksheets\sheet1.xml");
        }

        [Fact]
        public void ShouldRandomlyAddCells()
        {
            var testName = "ShouldRandomlyAddCells";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");

            ws.GetCell(10, 3);
            ws.GetCell("b", 1);
            ws.GetCell("C", 1);
            ws.GetCell("f6");
            ws.GetCell("A1");

            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}\{testName}.xml";

            TestHelper.AssertFile(expectedXmlFile, $@"{extractPath}\xl\worksheets\sheet1.xml");
        }
    }
}

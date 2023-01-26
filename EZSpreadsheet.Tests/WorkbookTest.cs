using Xunit.Abstractions;
using System.IO.Compression;
using DocumentFormat.OpenXml.Vml.Office;
using Xunit;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = Xunit.Assert;
using System.Text.RegularExpressions;

namespace EZSpreadsheet.Tests
{
    public class WorkbookTest
    {
        private readonly ITestOutputHelper output;

        public WorkbookTest(ITestOutputHelper output)
        {
            this.output = output;
            TestHelper.CreateFolder(TestHelper.TEST_OUTPUT_FOLDER);
        }        

        [Fact]
        public void ShouldGenerateWorkbook()
        {
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}\ShouldGenerateWorkbook.xlsx";
            var wb = new EZWorkbook(file);
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}\ShouldGenerateWorkbook";
            var extractedFiles = TestHelper.ExtractFiles(file, extractPath);

            var expectedFiles = new List<string>
            {
                $@"{extractPath}\xl\workbook.xml",
                $@"{extractPath}\xl\sharedStrings.xml",
                $@"{extractPath}\xl\styles.xml"
            };

            Assert.Equal(3, extractedFiles.Where(x => expectedFiles.Contains(x)).Count());

            var expectedXmlFolder = $@"{TestHelper.EXPECTED_XML_FOLDER}\ShouldGenerateWorkbook";

            TestHelper.AssertFile($@"{expectedXmlFolder}\workbook.xml", expectedFiles[0]);
            TestHelper.AssertFile($@"{expectedXmlFolder}\sharedStrings.xml", expectedFiles[1]);
            TestHelper.AssertFile($@"{expectedXmlFolder}\styles.xml", expectedFiles[2]);
        }

        [Fact]
        public void ShouldAddworksheets()
        {
            var testName = "ShouldAddworksheets";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}.xlsx";
            var wb = new EZWorkbook(file);
            wb.AddSheet("sheet1");
            wb.AddSheet("sheet2");
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}\{testName}";
            var extractedFiles = TestHelper.ExtractFiles(file, extractPath);

            var expectedFiles = new List<string>
            {
                $@"{extractPath}\xl\worksheets\sheet1.xml",
                $@"{extractPath}\xl\worksheets\sheet2.xml"
            };

            Assert.Equal(2, extractedFiles.Where(x => expectedFiles.Contains(x)).Count());

            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}\{testName}.xml";

            expectedFiles.ForEach(file => TestHelper.AssertFile(expectedXmlFile, file));            
        }

        [Fact]
        public void ShouldThrowExceptionWhenAddingSheetIfNameAlreadyExists()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldThrowExceptionWhenAddingSheetIfNameAlreadyExists.xlsx");
            workbook.AddSheet("NewSheet");
            Assert.Throws<Exception>(() => workbook.AddSheet("NewSheet"));
        }

        [Fact]
        public void ShouldGetSheet()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldGetSheet.xlsx");
            workbook.AddSheet("NewSheet");
            Assert.NotNull(workbook.GetSheet("NewSheet"));
        }

        [Fact]
        public void ShouldReturnNullIfSheetDoesNotExist()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldReturnNullIfSheetDoesNotExist.xlsx");
            workbook.AddSheet("NewSheet");
            Assert.Null(workbook.GetSheet("OldSheet"));
        }

        [Fact]
        public void ShouldGetSheetCount()
        {
            var workbook = new EZWorkbook($"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldGetSheetCount.xlsx");
            workbook.AddSheet("Sheet1");
            workbook.AddSheet("Sheet2");
            workbook.AddSheet("Sheet3");
            Assert.Equal(3, workbook.GetSheetCount());
        }
    }
}
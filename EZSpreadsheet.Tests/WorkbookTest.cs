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

            TestHelper.AssertFiles($@"{expectedXmlFolder}\workbook.xml", expectedFiles[0]);
            TestHelper.AssertFiles($@"{expectedXmlFolder}\sharedStrings.xml", expectedFiles[1]);
            TestHelper.AssertFiles($@"{expectedXmlFolder}\styles.xml", expectedFiles[2]);
        }
    }
}
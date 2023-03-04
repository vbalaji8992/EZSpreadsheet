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
        }        

        [Fact]
        public void ShouldGenerateWorkbook()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            wb.Save();         

            var expectedXmlFolder = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldGenerateWorkbook";
            TestHelper.AssertXml($@"{expectedXmlFolder}/workbook.xml", "xl/workbook.xml", memoryStream);
            TestHelper.AssertXml($@"{expectedXmlFolder}/sharedStrings.xml", "xl/sharedStrings.xml", memoryStream);
            TestHelper.AssertXml($@"{expectedXmlFolder}/styles.xml", "xl/styles.xml", memoryStream);
        }

        [Fact]
        public void ShouldAddworksheets()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            wb.AddSheet("sheet1");
            wb.AddSheet("sheet2");
            wb.Save();

            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldAddworksheets.xml";
            TestHelper.AssertXml(expectedXmlFile, "xl/worksheets/sheet1.xml", memoryStream);
            TestHelper.AssertXml(expectedXmlFile, "xl/worksheets/sheet2.xml", memoryStream);
        }

        [Fact]
        public void ShouldThrowExceptionWhenAddingSheetIfNameAlreadyExists()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            workbook.AddSheet("NewSheet");
            Assert.Throws<Exception>(() => workbook.AddSheet("NewSheet"));
        }

        [Fact]
        public void ShouldGetSheet()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            workbook.AddSheet("NewSheet");
            Assert.NotNull(workbook.GetSheet("NewSheet"));
        }

        [Fact]
        public void ShouldReturnNullIfSheetDoesNotExist()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            workbook.AddSheet("NewSheet");
            Assert.Null(workbook.GetSheet("OldSheet"));
        }

        [Fact]
        public void ShouldGetSheetCount()
        {
            var workbook = new EZWorkbook(new MemoryStream());
            workbook.AddSheet("Sheet1");
            workbook.AddSheet("Sheet2");
            workbook.AddSheet("Sheet3");
            Assert.Equal(3, workbook.GetSheetCount());
        }
    }
}
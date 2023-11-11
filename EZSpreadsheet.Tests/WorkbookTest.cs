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

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldGenerateWorkbook.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldAddworksheets()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            wb.AddSheet("sheet1");
            wb.AddSheet("sheet2");
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldAddworksheets.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
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
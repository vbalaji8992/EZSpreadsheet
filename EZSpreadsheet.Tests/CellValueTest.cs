using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet.Tests
{
    public class CellValueTest
    {
        public CellValueTest()
        {
        }

        [Fact]
        public void ShouldSetValueAsStringInCell1() 
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1");
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetValueAsStringInCell.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldSetValueAsStringInCell2()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue<string?>("Text1");
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetValueAsStringInCell.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldNotDuplicateValuesInSharedString()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1");
            ws.GetCell(1, 2).SetValue<string?>("Text1");
            ws.GetCell(2, 2).SetValue("Text1");
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldNotDuplicateValuesInSharedString.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldSetIntegerValues()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue(1);
            ws.GetCell(1, 2).SetValue<int?>(2);
            ws.GetCell(1, 3).SetValue(12.34);
            ws.GetCell(1, 4).SetValue<double?>(34.56);
            ws.GetCell(1, 5).SetValue(56.78f);
            ws.GetCell(1, 6).SetValue<float?>(78.90f);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetIntegerValues.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldConvertTextToNumber()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("1").ConvertToNumber();
            ws.GetCell(1, 2).SetValue("12.34").ConvertToNumber();
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldConvertTextToNumber.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldNotThrowExceptionForStringToNumberConversion()
        {
            var wb = new EZWorkbook(new MemoryStream());
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("text").ConvertToNumber();
            wb.Save();
        }
    }
}

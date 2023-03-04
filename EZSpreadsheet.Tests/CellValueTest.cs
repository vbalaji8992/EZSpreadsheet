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

            var expectedXmlFolder = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldSetValueAsStringInCell";
            TestHelper.AssertXml($"{expectedXmlFolder}/sheet1.xml", "xl/worksheets/sheet1.xml", memoryStream);            
            TestHelper.AssertXml($"{expectedXmlFolder}/sharedStrings.xml", "xl/sharedStrings.xml", memoryStream);            
        }

        [Fact]
        public void ShouldSetValueAsStringInCell2()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue<string?>("Text1");
            wb.Save();

            var expectedXmlFolder = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldSetValueAsStringInCell";
            TestHelper.AssertXml($"{expectedXmlFolder}/sheet1.xml", "xl/worksheets/sheet1.xml", memoryStream);
            TestHelper.AssertXml($"{expectedXmlFolder}/sharedStrings.xml", "xl/sharedStrings.xml", memoryStream);
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

            var expectedXmlFolder = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldNotDuplicateValuesInSharedString";
            TestHelper.AssertXml($"{expectedXmlFolder}/sheet1.xml", "xl/worksheets/sheet1.xml", memoryStream);
            TestHelper.AssertXml($"{expectedXmlFolder}/sharedStrings.xml", "xl/sharedStrings.xml", memoryStream);
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

            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldSetIntegerValues.xml";
            TestHelper.AssertXml(expectedXmlFile, "xl/worksheets/sheet1.xml", memoryStream);
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

            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}/ShouldConvertTextToNumber.xml";
            TestHelper.AssertXml(expectedXmlFile, "xl/worksheets/sheet1.xml", memoryStream);
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

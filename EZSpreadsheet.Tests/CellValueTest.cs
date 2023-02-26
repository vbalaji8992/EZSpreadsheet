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
            TestHelper.CreateFolder(TestHelper.TEST_OUTPUT_FOLDER);
        }

        [Fact]
        public void ShouldSetValueAsStringInCell1() 
        {
            var testName = "ShouldSetValueAsStringInCell";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1");
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile1 = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}/sheet1.xml";
            var expectedXmlFile2 = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}/sharedStrings.xml";

            TestHelper.AssertFile(expectedXmlFile1, $@"{extractPath}/xl/worksheets/sheet1.xml");
            TestHelper.AssertFile(expectedXmlFile2, $@"{extractPath}/xl/sharedStrings.xml");
        }

        [Fact]
        public void ShouldSetValueAsStringInCell2()
        {
            var testName = "ShouldSetValueAsStringInCell";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue<string?>("Text1");
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile1 = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}/sheet1.xml";
            var expectedXmlFile2 = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}/sharedStrings.xml";

            TestHelper.AssertFile(expectedXmlFile1, $@"{extractPath}/xl/worksheets/sheet1.xml");
            TestHelper.AssertFile(expectedXmlFile2, $@"{extractPath}/xl/sharedStrings.xml");
        }

        [Fact]
        public void ShouldNotDuplicateValuesInSharedString()
        {
            var testName = "ShouldNotDuplicateValuesInSharedString";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1");
            ws.GetCell(1, 2).SetValue<string?>("Text1");
            ws.GetCell(2, 2).SetValue("Text1");
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile1 = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}/sheet1.xml";
            var expectedXmlFile2 = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}/sharedStrings.xml";

            TestHelper.AssertFile(expectedXmlFile1, $@"{extractPath}/xl/worksheets/sheet1.xml");
            TestHelper.AssertFile(expectedXmlFile2, $@"{extractPath}/xl/sharedStrings.xml");
        }

        [Fact]
        public void ShouldSetIntegerValues()
        {
            var testName = "ShouldSetIntegerValues";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue(1);
            ws.GetCell(1, 2).SetValue<int?>(2);
            ws.GetCell(1, 3).SetValue(12.34);
            ws.GetCell(1, 4).SetValue<double?>(34.56);
            ws.GetCell(1, 5).SetValue(56.78f);
            ws.GetCell(1, 6).SetValue<float?>(78.90f);
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}.xml";

            TestHelper.AssertFile(expectedXmlFile, $@"{extractPath}/xl/worksheets/sheet1.xml");
        }

        [Fact]
        public void ShouldConvertTextToNumber()
        {
            var testName = "ShouldConvertTextToNumber";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("1").ConvertToNumber();
            ws.GetCell(1, 2).SetValue("12.34").ConvertToNumber();
            wb.Save();

            var extractPath = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}";
            TestHelper.ExtractFiles(file, extractPath);
            var expectedXmlFile = $@"{TestHelper.EXPECTED_XML_FOLDER}/{testName}.xml";

            TestHelper.AssertFile(expectedXmlFile, $@"{extractPath}/xl/worksheets/sheet1.xml");
        }

        [Fact]
        public void ShouldNotThrowExceptionForStringToNumberConversion()
        {
            var testName = "ShouldNotThrowExceptionForStringToNumberConversion";
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/{testName}.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("text").ConvertToNumber();
            wb.Save();
        }
    }
}

using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System;
using System.Collections.Generic;
using System.IO;
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

        [Fact]
        public void ShouldInsert1DListOfStrings()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            var list = new List<string>();
            for (uint i = 1; i <= 10; i++)
                list.Add("Text-" + i.ToString());

            ws.GetCell(2, 2).InsertData(list);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldInsert1DListOfStrings.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldInsert1DListOfNumbers()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            var list = new List<uint>();
            for (uint i = 1; i <= 10; i++)
                list.Add(i);

            ws.GetCell(2, 2).InsertData(list);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldInsert1DListOfNumbers.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldInsert2DList()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            var list = new List<Tuple<string, double, int>>()
            { 
                Tuple.Create("Jack", 78.8, 8),
                Tuple.Create("Abbey", 92.1, 9),
                Tuple.Create("Dave", 88.3, 9),
                Tuple.Create("Sam", 91.7, 8),
                Tuple.Create("Ed", 71.2, 5),
                Tuple.Create("Penelope", 82.9, 8),
                Tuple.Create("Linda", 99.0, 9),
                Tuple.Create("Judith", 84.3, 9) 
            };

            ws.GetCell(2, 2).InsertData(list);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldInsert2DList.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldInsert2DListWithPropertyNameAsHeading()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            var list = new List<Tuple<string, double, int>>()
            {
                Tuple.Create("Jack", 78.8, 8),
                Tuple.Create("Abbey", 92.1, 9),
                Tuple.Create("Dave", 88.3, 9),
                Tuple.Create("Sam", 91.7, 8),
                Tuple.Create("Ed", 71.2, 5),
                Tuple.Create("Penelope", 82.9, 8),
                Tuple.Create("Linda", 99.0, 9),
                Tuple.Create("Judith", 84.3, 9)
            };

            ws.GetCell(2, 2).InsertData(list, true);
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldInsert2DListWithPropertyNameAsHeading.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }

        [Fact]
        public void ShouldSetCellFormula()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");

            var list = new List<uint>();
            for (uint i = 1; i <= 10; i++)
                list.Add(i);

            ws.GetCell(2, 2).InsertData(list);
            ws.GetCell("d6").SetFormula("SUM(B2:B11)");
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetCellFormula.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }
    }
}

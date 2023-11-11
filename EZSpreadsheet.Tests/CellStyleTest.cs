using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet.Tests
{
    public class CellStyleTest
    {
        public CellStyleTest()
        {

        }

        [Fact]
        public void ShouldSetFontToCell()
        {
            var memoryStream = new MemoryStream();
            var wb = new EZWorkbook(memoryStream);
            var ws = wb.AddSheet("sheet1");
            ws.GetCell(1, 1).SetValue("Text1").SetStyle(new Style.EZStyle { Font = Style.EZFont.Arial });
            ws.GetCell(2, 1).SetValue("Text2").SetStyle(new Style.EZStyle { Font = Style.EZFont.Calibri });
            wb.Save();

            var expectedFile = $@"{TestHelper.EXPECTED_FILES_FOLDER}/ShouldSetFontToCell.xlsx";
            TestHelper.AssertSpreadsheet(memoryStream, expectedFile);
        }
    }
}

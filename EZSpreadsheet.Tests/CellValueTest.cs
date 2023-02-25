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
        public void ShouldSetTextInACell() 
        {
            var file = $@"{TestHelper.TEST_OUTPUT_FOLDER}/ShouldSetTextInACell.xlsx";
            var wb = new EZWorkbook(file);
            var ws = wb.AddSheet("sheet1");
            wb.Save();
        }
    }
}

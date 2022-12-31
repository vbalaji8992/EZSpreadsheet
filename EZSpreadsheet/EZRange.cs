using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    public class EZRange
    {
        public EZWorksheet Worksheet { get; }
        public List<EZCell> CellList { get; }

        public EZRange(EZWorksheet worksheet, List<EZCell> cellList)
        {
            Worksheet = worksheet;
            CellList = cellList;
        }

        public void ConvertToNumber()
        {
            foreach(var cell in CellList)
            {
                cell.ConvertToNumber();
            }
        }

        public void SetFontStyle(EZFontStyle fontStyle)
        {
            foreach (var cell in CellList)
            {
                cell.SetFontStyle(fontStyle);
            }
        }
    }
}

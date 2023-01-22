using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EZSpreadsheet.Style;

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

        public EZRange ConvertToNumber()
        {
            foreach(var cell in CellList)
            {
                cell.ConvertToNumber();
            }

            return this;
        }

        public EZRange SetStyle(EZStyle cellStyle)
        {
            foreach (var cell in CellList)
            {
                cell.SetStyle(cellStyle);
            }

            return this;
        }
    }
}

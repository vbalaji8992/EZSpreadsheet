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
        public EZCell StartCell { get; }
        public EZCell EndCell { get; }

        public EZRange(EZWorksheet worksheet, EZCell startCell, EZCell endCell)
        {
            Worksheet = worksheet;
            StartCell = startCell;
            EndCell = endCell;
        }      
    }
}

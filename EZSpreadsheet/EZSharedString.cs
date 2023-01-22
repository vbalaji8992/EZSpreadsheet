using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    internal class EZSharedString
    {
        internal EZWorkbook WorkBook { get; }

        internal SharedStringTablePart SharedStringPart { get; }

        internal Dictionary<string, int> StringTable { get; }

        internal EZSharedString(EZWorkbook workBook, SharedStringTablePart sharedStringTablePart)
        {
            WorkBook = workBook;
            SharedStringPart = sharedStringTablePart;
            SharedStringPart.SharedStringTable = new SharedStringTable();
            StringTable = new Dictionary<string, int>();
        }

        internal int InsertString(string value)
        {
            if (SharedStringPart.SharedStringTable == null)
            {
                SharedStringPart.SharedStringTable = new SharedStringTable();
            }                      

            if (!StringTable.ContainsKey(value))
            {
                SharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(value)));
                StringTable.Add(value, StringTable.Count);
            }

            return StringTable[value];
        }
    }
}

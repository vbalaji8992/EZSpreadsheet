using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    class EZSharedString
    {
        public EZWorkbook WorkBook { get; }

        public SharedStringTablePart SharedStringPart { get; }

        public Dictionary<string, int> StringTable { get; }

        public EZSharedString(EZWorkbook workBook, SharedStringTablePart sharedStringTablePart)
        {
            WorkBook = workBook;
            SharedStringPart = sharedStringTablePart;
            StringTable = new Dictionary<string, int>();
        }

        public int InsertString(string value)
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

            //SharedStringPart.SharedStringTable.Save();

            return StringTable[value];
        }
    }
}

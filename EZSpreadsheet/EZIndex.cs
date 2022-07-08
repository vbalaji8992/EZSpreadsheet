using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    class EZIndex
    {
        public const int MaxColumnIndex = 16384;
        public const uint AsciiOffset = 64;

        public static (string columnName, uint rowIndex) GetRowColumnIndex(string cellReference)
        {
            if (Regex.IsMatch(cellReference, "^[a-zA-Z]{1,2}[0-9]+$"))
            {
                string columnName = Regex.Replace(cellReference, @"[\d]", "");
                uint rowIndex = Convert.ToUInt32(Regex.Replace(cellReference, "[a-zA-Z]", ""));
                return (columnName, rowIndex);
            }
            else
            {
                throw new ArgumentOutOfRangeException("Invalid cell reference");
            }
        }

        public static string GetColumnName(uint columnIndex)
        {
            if (columnIndex > MaxColumnIndex)
            {
                throw new ArgumentOutOfRangeException("Invalid column index");
            }

            if (columnIndex <= 26)
            {
                uint columnUnicode = columnIndex + 64;

                return ((char)columnUnicode).ToString();
            }

            uint num = columnIndex / 26;
            uint rem = columnIndex % 26;

            if (rem == 0)
            {
                num -= 1;
                rem = 26;
            }

            return GetColumnName(num) + ((char)(rem + AsciiOffset)).ToString();
        }

        public static uint GetColumnIndex(string columnName)
        {
            if (columnName.Length == 1)
            {
                return columnName.ToUpper()[0] - AsciiOffset;
            }

            return GetColumnIndex(columnName.Substring(0, columnName.Length - 1)) * 26 + (columnName.ToUpper()[columnName.Length - 1] - AsciiOffset);
        }
    }
}

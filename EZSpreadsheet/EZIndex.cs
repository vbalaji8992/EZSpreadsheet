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
        internal const int MaxColumnIndex = 16384;
        internal const uint AsciiOffset = 64;

        internal static (string columnName, uint rowIndex) GetRowIndexColumnName(string cellReference)
        {
            string columnName;
            uint rowIndex;

            if (Regex.IsMatch(cellReference, "^[a-zA-Z]{1,2}[0-9]+$"))
            {
                columnName = Regex.Replace(cellReference, @"[\d]", "").ToUpper();
                rowIndex = Convert.ToUInt32(Regex.Replace(cellReference, "[a-zA-Z]", ""));                
            }
            else
            {
                throw new ArgumentOutOfRangeException("Invalid cell reference");
            }

            CheckForInvalidIndex(columnName, rowIndex);

            return (columnName, rowIndex);
        }

        internal static string GetColumnName(uint columnIndex)
        {
            if (columnIndex < 1 || columnIndex > MaxColumnIndex)
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

        internal static uint GetColumnIndex(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentOutOfRangeException("Invalid column name");
            }

            if (columnName.Length == 1)
            {
                return columnName.ToUpper()[0] - AsciiOffset;
            }

            return GetColumnIndex(columnName.Substring(0, columnName.Length - 1)) * 26 + (columnName.ToUpper()[columnName.Length - 1] - AsciiOffset);
        }

        internal static void CheckForInvalidIndex(string columnName, uint rowIndex)
        {
            uint columnIndex = GetColumnIndex(columnName);

            if (columnIndex > MaxColumnIndex)
            {
                throw new ArgumentOutOfRangeException("Invalid column name");
            }

            if (rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException("Invalid row index");
            }
        }
    }
}

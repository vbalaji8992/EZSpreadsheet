using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    public class EZCell
    {
        public EZWorksheet Worksheet { get; }
        public uint RowIndex { get; set; }
        public string ColumnName { get; set; }
        public uint ColumnIndex { get; set; }
        public string CellReference { get { return ColumnName + RowIndex; } }
        public Cell Cell { get; }
        public Row Row { get; }

        public EZCell(uint rowIndex, string columnName, Row row, Cell cell, EZWorksheet worksheet)
        {
            RowIndex = rowIndex;
            ColumnName = columnName;
            Worksheet = worksheet;
            Row = row;
            Cell = cell;
            ColumnIndex = EZIndex.GetColumnIndex(columnName);
        }

        public void SetText<T>(T value)
        {
            int index = Worksheet.WorkBook.SharedString.InsertString(value.ToString());

            Cell.CellValue = new CellValue(index.ToString());
            Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            
            //Worksheet.SaveWorksheet();
        }

        public void SetValue(int value)
        {
            Cell.CellValue = new CellValue(value.ToString());
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }

        public void ApplyFormat()
        {
            Cell.StyleIndex = 1;
        }

    }
}

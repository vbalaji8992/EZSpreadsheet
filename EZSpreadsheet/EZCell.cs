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

        public void ApplyStyle(uint index)
        {
            Cell.StyleIndex = index;
        }

        public void ConvertToNumber()
        {
            if (Cell.DataType == null || Cell.CellValue == null || Cell.DataType == CellValues.Number)
            {
                return;
            }

            int indexInStringTable;
            try
            {
                indexInStringTable = Convert.ToInt32(Cell.CellValue.InnerText);
            }
            catch (Exception)
            {
                throw new ArgumentException("Not possible to convert to a number");
            }

            var kvp = Worksheet.WorkBook.SharedString.StringTable.First(x => x.Value == indexInStringTable);

            if (kvp.Key != null)
            {
                SetValue(Convert.ToInt32(kvp.Key));
            }
        }

        public EZRange? InsertData<T>(List<T> data, bool includePropNameAsHeading = false)
        {
            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {
                return InsertValueType(data);                
            }

            uint currentRow = RowIndex;

            var props = typeof(T).GetProperties();

            if (includePropNameAsHeading)
            {
                var propNames = props.Select(prop => prop.Name).ToList();
                InsertValueType(propNames, true);
                currentRow++;
            }

            var firstCell = Worksheet.GetCell(currentRow, ColumnIndex);
            var currentCell = firstCell;       

            foreach (var item in data)
            {
                uint currentColumn = ColumnIndex;

                foreach (var prop in props)
                {
                    var value = item?.GetType().GetProperty(prop.Name)?.GetValue(item)?.ToString() ?? "";
                    currentCell = Worksheet.GetCell(currentRow, currentColumn);
                    currentCell.SetText(value);
                    currentColumn++;
                }

                currentRow++;
            }

            return new EZRange(Worksheet, firstCell, currentCell);
        }

        internal EZRange InsertValueType<T>(List<T> data, bool transposeData = false)
        {
            var currentRow = RowIndex;
            var currentColumn = ColumnIndex;
            var currentCell = this;

            foreach (var value in data)
            {
                currentCell = Worksheet.GetCell(currentRow, currentColumn);
                currentCell.SetText(value);

                if (transposeData)
                {
                    currentColumn++;
                }
                else
                {
                    currentRow++;
                }
            }

            return new EZRange(Worksheet, this, currentCell);
        }
    }
}

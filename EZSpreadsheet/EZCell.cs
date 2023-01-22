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

        private static HashSet<Type> IntegralTypes = new HashSet<Type>
        {
            typeof(sbyte),
            typeof(byte),
            typeof(short),
            typeof(ushort),
            typeof(int),
            typeof(uint),
            typeof(long),
            typeof(ulong)
        };

        private static HashSet<Type> DecimalTypes = new HashSet<Type>
        {
            typeof(float),
            typeof(double),
            typeof(decimal)
        };

        public EZCell(uint rowIndex, string columnName, Row row, Cell cell, EZWorksheet worksheet)
        {
            RowIndex = rowIndex;
            ColumnName = columnName;
            Worksheet = worksheet;
            Row = row;
            Cell = cell;
            ColumnIndex = EZIndex.GetColumnIndex(columnName);
        }

        public EZCell SetValue<T>(T value)
        {
            if (value != null && isNumericType(typeof(T)))
            {
                SetNumber(value.ToString()!);
            }
            else
            {
                int index = Worksheet.WorkBook.SharedString.InsertString(value?.ToString() ?? "null");
                Cell.CellValue = new CellValue(index.ToString());
                Cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }

            return this;
        }

        private bool isNumericType(Type type) 
        {
            var nullableType = Nullable.GetUnderlyingType(type);
            if (nullableType != null)
                type = nullableType;

            if (IntegralTypes.Contains(type) || DecimalTypes.Contains(type))
                return true;

            return false;
        }

        private void SetNumber(string value)
        { 
            Cell.CellValue = new CellValue(value);
            Cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }

        public EZCell SetFormula(string formula)
        {
            Cell.CellFormula = new CellFormula(formula);
            return this;
        }

        public void ApplyStyle(uint index)
        {
            Cell.StyleIndex = index;
        }

        public EZCell ConvertToNumber()
        {
            if (Cell.DataType == null || Cell.CellValue == null || Cell.DataType == CellValues.Number)
            {
                return this;
            }

            int indexInStringTable;
            try
            {
                indexInStringTable = Convert.ToInt32(Cell.CellValue.InnerText);
            }
            catch (Exception)
            {
                return this;
            }

            var kvp = Worksheet.WorkBook.SharedString.StringTable.First(x => x.Value == indexInStringTable);

            if (kvp.Key != null)
            {
                try
                {
                    SetValue(Convert.ToInt32(kvp.Key));
                }
                catch (Exception)
                {
                    return this;
                }
            }

            return this;
        }

        public EZRange InsertData<T>(List<T> data, bool includePropNameAsHeading = false)
        {
            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {
                return InsertValueType(data);                
            }

            uint currentRow = RowIndex;

            var props = typeof(T).GetProperties();

            EZRange range = new EZRange(Worksheet, new List<EZCell>());

            if (includePropNameAsHeading)
            {
                var propNames = props.Select(prop => prop.Name).ToList();
                range = InsertValueType(propNames, true);
                currentRow++;
            }

            foreach (var item in data)
            {
                uint currentColumn = ColumnIndex;

                foreach (var prop in props)
                {
                    var propInfo = item?.GetType().GetProperty(prop.Name);
                    var type = propInfo?.PropertyType;
                    var value = propInfo?.GetValue(item);                    
                    var currentCell = Worksheet.GetCell(currentRow, currentColumn);

                    if (value != null && type != null)
                        currentCell.SetValue(CastHelper.Cast(value, type));
                    else
                        currentCell.SetValue("null");

                    range.CellList.Add(currentCell);
                    currentColumn++;
                }

                currentRow++;
            }

            return range;
        }

        internal EZRange InsertValueType<T>(List<T> data, bool transposeData = false)
        {
            var currentRow = RowIndex;
            var currentColumn = ColumnIndex;

            var cellList = new List<EZCell>();

            foreach (var value in data)
            {
                var currentCell = Worksheet.GetCell(currentRow, currentColumn);
                currentCell.SetValue(value);
                cellList.Add(currentCell);

                if (transposeData)
                {
                    currentColumn++;
                }
                else
                {
                    currentRow++;
                }
            }

            return new EZRange(Worksheet, cellList);
        }

        public EZCell SetStyle(EZCellStyle cellStyle)
        {
            var style = Worksheet.WorkBook.StyleSheet.AppendCellStyle(cellStyle);
            var styleIndex = Worksheet.WorkBook.StyleSheet.AppendCellFormat(style);
            ApplyStyle(styleIndex);
            return this;
        }
    }
}

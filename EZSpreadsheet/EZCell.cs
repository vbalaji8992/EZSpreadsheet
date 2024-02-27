using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using EZSpreadsheet.Style;
using EZSpreadsheet.Utils;

namespace EZSpreadsheet
{
    public class EZCell
    {
        internal EZWorksheet Worksheet { get; }
        internal uint RowIndex { get; set; }
        internal string ColumnName { get; set; }
        internal uint ColumnIndex { get; set; }
        internal string CellReference { get { return ColumnName + RowIndex; } }
        internal Cell Cell { get; }
        internal Row Row { get; }

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

        internal EZCell(uint rowIndex, string columnName, Row row, Cell cell, EZWorksheet worksheet)
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

        private void ApplyStyle(uint index)
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
                    SetValue(Convert.ToDouble(kvp.Key));
                }
                catch (Exception)
                {
                    return this;
                }
            }

            return this;
        }

        public EZRange InsertData<T>(IEnumerable<T> data, EZListOptions? listOptions = null)
        {
            if (listOptions == null)
                listOptions = new EZListOptions();

            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {
                return InsertValueType(data, listOptions!.TransposeData);                
            }

            uint currentRow = RowIndex;
            uint currentColumn = ColumnIndex;

            var props = typeof(T).GetProperties();

            EZRange range = new EZRange(Worksheet, new List<EZCell>());

            bool doNotTranspose = !listOptions!.TransposeData;

            if (listOptions!.AddPropertyNameAsHeading)
            {
                var propNames = props.Select(prop => prop.Name).ToList();
                range = InsertValueType(propNames, doNotTranspose);
                if (doNotTranspose)
                    currentRow++;
                else
                    currentColumn++;
            }

            foreach (var item in data)
            {
                if (doNotTranspose)
                    currentColumn = ColumnIndex;
                else
                    currentRow = RowIndex;

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

                    if (doNotTranspose)
                        currentColumn++;
                    else
                        currentRow++;
                }

                if (doNotTranspose)
                    currentRow++;
                else
                    currentColumn++;
            }

            return range;
        }

        private EZRange InsertValueType<T>(IEnumerable<T> data, bool transposeData)
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
                    currentColumn++;
                else
                    currentRow++;
            }

            return new EZRange(Worksheet, cellList);
        }

        public EZCell SetStyle(EZStyle cellStyle)
        {
            var style = Worksheet.WorkBook.StyleSheet.AppendCellStyle(cellStyle);
            var styleIndex = Worksheet.WorkBook.StyleSheet.AppendCellFormat(style);
            ApplyStyle(styleIndex);
            return this;
        }
    }
}

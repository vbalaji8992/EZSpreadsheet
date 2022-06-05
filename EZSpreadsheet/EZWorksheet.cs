using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EZSpreadsheet
{
    public class EZWorksheet
    {
        public EZWorkbook WorkBook { get; }
        public Worksheet Worksheet { get; set; }
        public Sheet Sheet { get; set; }
        public SheetData SheetData { get; set; }
        public Dictionary<uint, List<EZCell>> CellsKvp { get; }

        public EZWorksheet(Worksheet worksheet, Sheet sheet, EZWorkbook workBook)
        {
            WorkBook = workBook;
            Worksheet = worksheet;
            Sheet = sheet;
            SheetData = worksheet.GetFirstChild<SheetData>();
            CellsKvp = new Dictionary<uint, List<EZCell>>();
        }

        public EZCell GetCell(string columnName, uint rowIndex)
        {
            if (EZIndex.GetColumnIndex(columnName) > EZIndex.MaxColumnIndex || rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException("Invalid column name");
            }

            EZCell excelCell = null;

            if (CellsKvp.ContainsKey(rowIndex))
            {
                excelCell = CellsKvp[rowIndex].Where(x => x.ColumnName == columnName && x.RowIndex == rowIndex).FirstOrDefault();
            }

            if (excelCell != null)
            {
                return excelCell;
            }

            var cell = AddCell(columnName, rowIndex);

            return cell;
        }

        public EZCell GetCell(uint rowIndex, uint columnIndex)
        {
            string columnName = EZIndex.GetColumnName(columnIndex);

            return GetCell(columnName, rowIndex);
        }
        
        public EZCell GetCell(string cellReference)
        {
            var (columnName, rowIndex) = EZIndex.GetRowColumnIndex(cellReference);

            return GetCell(columnName, rowIndex);
        }

        private EZCell AddCell(string columnName, uint rowIndex)
        {
            string cellReference = columnName + rowIndex;

            if (!CellsKvp.ContainsKey(rowIndex))
            {
                Row row = new Row() { RowIndex = rowIndex };
                SheetData.Append(row);

                Cell newCell = new Cell() { CellReference = cellReference };
                row.Append(newCell);

                EZCell excelCell = new EZCell(rowIndex, columnName, row, newCell, this);
                CellsKvp.Add(rowIndex, new List<EZCell> { excelCell });

                return excelCell;
            }
            else
            {
                var cellsWithRowIndex = CellsKvp[rowIndex];
                EZCell refCell = null;

                foreach (var cell in cellsWithRowIndex)
                {
                    if (string.Compare(cell.CellReference, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Row row;
                Cell newCell = new Cell() { CellReference = cellReference };

                if (refCell == null)
                {
                    row = cellsWithRowIndex.First().Row;
                    row.Append(newCell);
                }
                else
                {
                    row = refCell.Row;
                    row.InsertBefore(newCell, refCell.Cell);
                }

                EZCell excelCell = new EZCell(rowIndex, columnName, row, newCell, this);
                CellsKvp[rowIndex].Add(excelCell);

                return excelCell;
            }
        }

        public void InsertData<T>(List<T> data, string cellReference)
        {
            var (columnName, rowIndex) = EZIndex.GetRowColumnIndex(cellReference);
            var startRowIndex = rowIndex;
            var startColumnIndex = EZIndex.GetColumnIndex(columnName);

            uint currentRow = startRowIndex;

            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {
                foreach (var value in data)
                {
                    GetCell(currentRow, startColumnIndex).SetText(value);
                    currentRow++;
                }

                return;
            }

            var props = typeof(T).GetProperties();            

            foreach (var item in data)
            {
                uint currentColumn = startColumnIndex;

                foreach (var prop in props)
                {
                    var value = item.GetType().GetProperty(prop.Name).GetValue(item).ToString();
                    GetCell(currentRow, currentColumn).SetText(value);
                    currentColumn++;
                }

                currentRow++;
            }
        }

        public void SaveWorksheet()
        {
            Worksheet.Save();
        }
    }
}

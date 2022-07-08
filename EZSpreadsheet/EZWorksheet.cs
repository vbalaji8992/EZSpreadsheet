using DocumentFormat.OpenXml.Packaging;
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
        internal EZWorkbook WorkBook { get; }
        internal Worksheet Worksheet { get; set; }
        internal Sheet Sheet { get; set; }
        internal SheetData SheetData { get; set; }
        internal Dictionary<uint, List<EZCell>> CellListByRowIndex { get; }

        public EZWorksheet(EZWorkbook workBook, string? sheetName)
        {
            WorksheetPart worksheetPart = workBook.SpreadsheetDocument.WorkbookPart!.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            Sheet sheet = new Sheet()
            {
                Id = workBook.SpreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = workBook.NextAvailableSheetId,
                Name = (string.IsNullOrEmpty(sheetName)) ? "Sheet" + workBook.NextAvailableSheetId : sheetName
            };

            workBook.Sheets.Append(sheet);

            WorkBook = workBook;
            Worksheet = worksheetPart.Worksheet;
            Sheet = sheet;
            SheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            CellListByRowIndex = new Dictionary<uint, List<EZCell>>();
        }

        public EZCell GetCell(string columnName, uint rowIndex)
        {
            if (EZIndex.GetColumnIndex(columnName) > EZIndex.MaxColumnIndex || rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException("Invalid column name");
            }

            EZCell? excelCell = null;

            if (CellListByRowIndex.ContainsKey(rowIndex))
            {
                excelCell = CellListByRowIndex[rowIndex].Where(x => x.ColumnName == columnName && x.RowIndex == rowIndex).FirstOrDefault();
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

            if (!CellListByRowIndex.ContainsKey(rowIndex))
            {
                Row row = new Row() { RowIndex = rowIndex };
                SheetData.Append(row);

                Cell newCell = new Cell() { CellReference = cellReference };
                row.Append(newCell);

                EZCell excelCell = new EZCell(rowIndex, columnName, row, newCell, this);
                CellListByRowIndex.Add(rowIndex, new List<EZCell> { excelCell });

                return excelCell;
            }
            else
            {
                var cellsWithRowIndex = CellListByRowIndex[rowIndex];
                EZCell? refCell = null;

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
                CellListByRowIndex[rowIndex].Add(excelCell);

                return excelCell;
            }
        }

        public void InsertData<T>(List<T> data, string cellReference, bool includePropNameAsHeading = false)
        {           
            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {
                InsertValueType(data, cellReference);
                return;
            }

            var (columnName, rowIndex) = EZIndex.GetRowColumnIndex(cellReference);
            var startRowIndex = rowIndex;
            var startColumnIndex = EZIndex.GetColumnIndex(columnName);

            uint currentRow = startRowIndex;

            var props = typeof(T).GetProperties();
            
            if (includePropNameAsHeading)
            {
                var propNames = props.Select(prop => prop.Name).ToList();
                InsertValueType(propNames, cellReference, true);
                currentRow++;
            }

            foreach (var item in data)
            {
                uint currentColumn = startColumnIndex;

                foreach (var prop in props)
                {
                    var value = item?.GetType().GetProperty(prop.Name)?.GetValue(item)?.ToString();
                    GetCell(currentRow, currentColumn).SetText(value);
                    currentColumn++;
                }

                currentRow++;
            }
        }

        public void InsertValueType<T>(List<T> data, string cellReference, bool transposeData = false)
        {
            var (columnName, rowIndex) = EZIndex.GetRowColumnIndex(cellReference);
            var currentRow = rowIndex;
            var currentColumn = EZIndex.GetColumnIndex(columnName);

            foreach (var value in data)
            {
                GetCell(currentRow, currentColumn).SetText(value);

                if (transposeData)
                {
                    currentColumn++;
                }
                else
                {
                    currentRow++;
                }                
            }
        }

        public uint GetFirstRowIndex()
        {
            Row? firstRow = SheetData.FirstChild as Row;

            return firstRow?.RowIndex ?? throw new Exception("Empty sheet");
        }

        public uint GetLastRowIndex()
        {
            Row? lastRow = SheetData.LastChild as Row;

            return lastRow?.RowIndex ?? throw new Exception("Empty sheet");
        }

        public uint GetFirstColumnIndex()
        {
            uint firstColumnIndex = EZIndex.MaxColumnIndex;

            foreach (var kvp in CellListByRowIndex)
            {               
                foreach(var cell in kvp.Value)
                {
                    if (cell.ColumnIndex < firstColumnIndex)
                    {
                        firstColumnIndex = cell.ColumnIndex;
                    }
                }
            }

            if(firstColumnIndex == EZIndex.MaxColumnIndex)
            {
                throw new Exception("Empty sheet");
            }

            return firstColumnIndex;
        }

        public uint GetLastColumnIndex()
        {
            uint lastColumnIndex = 0;

            foreach (var kvp in CellListByRowIndex)
            {
                foreach (var cell in kvp.Value)
                {
                    if (cell.ColumnIndex > lastColumnIndex)
                    {
                        lastColumnIndex = cell.ColumnIndex;
                    }
                }
            }

            if (lastColumnIndex == 0)
            {
                throw new Exception("Empty sheet");
            }

            return lastColumnIndex;
        }

        public string GetFirstColumnName()
        {
            return EZIndex.GetColumnName(GetFirstColumnIndex());
        }

        public string GetLastColumnName()
        {
            return EZIndex.GetColumnName(GetLastColumnIndex());
        }

        public string GetSheetName()
        {
            return Sheet.Name!;
        }

        public void SaveWorksheet()
        {
            Worksheet.Save();
        }
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EZSpreadsheet.Utils;

namespace EZSpreadsheet
{
    public class EZWorksheet
    {
        internal EZWorkbook WorkBook { get; }
        internal Worksheet Worksheet { get; set; }
        internal Sheet Sheet { get; set; }
        internal SheetData SheetData { get; set; }
        internal Dictionary<uint, Dictionary<uint, EZCell>> CellsyRowColumnIndex { get; }

        internal EZWorksheet(EZWorkbook workBook, string? sheetName)
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
            CellsyRowColumnIndex = new Dictionary<uint, Dictionary<uint, EZCell>>();
        }

        public EZCell GetCell(string columnName, uint rowIndex)
        {
            columnName = columnName.ToUpper();
            var columnIndex = EZIndex.GetColumnIndex(columnName);

            EZIndex.CheckForInvalidIndex(columnName, rowIndex);

            EZCell? excelCell = null;

            if (CellsyRowColumnIndex.ContainsKey(rowIndex))
            {
                var rowCells = CellsyRowColumnIndex[rowIndex];
                if (rowCells.ContainsKey(columnIndex))
                    excelCell = rowCells[columnIndex];
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
            var (columnName, rowIndex) = EZIndex.GetRowIndexColumnName(cellReference);

            return GetCell(columnName, rowIndex);
        }

        private EZCell AddCell(string columnName, uint rowIndex)
        {
            if (!CellsyRowColumnIndex.ContainsKey(rowIndex))
            {
                return AddRowAndAppendCell(columnName, rowIndex);
            }
            else
            {
                return AppendCellToRow(columnName, rowIndex);
            }
        }

        private EZCell AppendCellToRow(string columnName, uint rowIndex)
        {
            string cellReference = columnName + rowIndex;
            var columnIndex = EZIndex.GetColumnIndex(columnName);

            var cellsWithRowIndex = CellsyRowColumnIndex[rowIndex];
            EZCell? refCell = null;

            var maxCol = cellsWithRowIndex.Keys.Max()!;
            if (columnIndex < maxCol)
            {
                foreach (var kvp in cellsWithRowIndex)
                {
                    if (kvp.Key > columnIndex)
                    {
                        refCell = kvp.Value;
                        break;
                    }
                }
            }

            Row row;
            Cell newCell = new Cell() { CellReference = cellReference };

            if (refCell == null)
            {
                row = cellsWithRowIndex.First().Value.Row;
                row.Append(newCell);
            }
            else
            {
                row = refCell.Row;
                row.InsertBefore(newCell, refCell.Cell);
            }

            EZCell excelCell = new EZCell(rowIndex, columnName, row, newCell, this);
            CellsyRowColumnIndex[rowIndex].Add(columnIndex, excelCell);

            return excelCell;
        }

        private EZCell AddRowAndAppendCell(string columnName, uint rowIndex)
        {
            string cellReference = columnName + rowIndex;
            var columnIndex = EZIndex.GetColumnIndex(columnName);

            Row row = new Row() { RowIndex = rowIndex };   

            var maxRow = (CellsyRowColumnIndex.Count > 0) ? CellsyRowColumnIndex.Keys.Max() : 0;
            Row? refRow = null;
            if (rowIndex < maxRow)
            {
                var presentRows = SheetData.ChildElements;
                refRow = presentRows?.Select(x => x as Row).Where(x => x?.RowIndex! > rowIndex).OrderBy(x => x?.RowIndex).FirstOrDefault();
            }            

            if (refRow != null)
            {
                SheetData.InsertBefore(row, refRow);
            }
            else
            {
                SheetData.Append(row);
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.Append(newCell);

            EZCell excelCell = new EZCell(rowIndex, columnName, row, newCell, this);
            CellsyRowColumnIndex.Add(rowIndex, new Dictionary<uint, EZCell> { { columnIndex, excelCell } });

            return excelCell;
        }

        public EZRange GetRange(uint startRowIndex, uint startColIndex, uint endRowIndex, uint endColIndex)
        {
            var cellList = new List<EZCell>();

            if (startRowIndex > endRowIndex)
            {
                uint temp = startRowIndex;
                startRowIndex = endRowIndex;
                endRowIndex = temp;
            }

            if (startColIndex > endColIndex)
            {
                uint temp = startColIndex;
                startColIndex = endColIndex;
                endColIndex = temp;
            }
            
            for (uint i = startRowIndex; i <= endRowIndex; i++)
            {
                for (uint j = startColIndex; j <= endColIndex; j++)
                {
                    cellList.Add(GetCell(i, j));
                }
            }

            return new EZRange(this, cellList);
        }

        public EZRange GetRange(string startCellReference, string endCellReference) 
        {
            var startRowColumn = EZIndex.GetRowIndexColumnName(startCellReference);
            var endRowColumn = EZIndex.GetRowIndexColumnName(endCellReference);

            return GetRange(
                startRowColumn.rowIndex,
                EZIndex.GetColumnIndex(startRowColumn.columnName),
                endRowColumn.rowIndex,
                EZIndex.GetColumnIndex(endRowColumn.columnName)
            );
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

            foreach (var kvp in CellsyRowColumnIndex)
            {               
                foreach(var cell in kvp.Value)
                {
                    if (cell.Value.ColumnIndex < firstColumnIndex)
                    {
                        firstColumnIndex = cell.Value.ColumnIndex;
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

            foreach (var kvp in CellsyRowColumnIndex)
            {
                foreach (var cell in kvp.Value)
                {
                    if (cell.Value.ColumnIndex > lastColumnIndex)
                    {
                        lastColumnIndex = cell.Value.ColumnIndex;
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

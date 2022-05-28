using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EZSpreadsheet
{
    public class EZWorkbook
    {
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly Sheets sheets;
        private uint nextAvailableSheetId = 1;

        internal List<EZWorksheet> Worksheets { get; }
        internal EZSharedString SharedString { get; }
        internal EZStyle Style { get; }

        public EZWorkbook(string filepath)
        {

            spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            spreadsheetDocument.AddWorkbookPart();
            spreadsheetDocument.WorkbookPart.Workbook = new Workbook();

            sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Worksheets = new List<EZWorksheet>();
            //AddSheet();

            SharedStringTablePart sharedStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            SharedString = new EZSharedString(this, sharedStringPart);

            WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            Style = new EZStyle(this, workbookStylesPart);
        }

        public EZWorksheet AddSheet(string sheetName = null)
        {
            WorksheetPart worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = nextAvailableSheetId,
                Name = (string.IsNullOrEmpty(sheetName)) ? "Sheet" + nextAvailableSheetId : sheetName
            };

            sheets.Append(sheet);

            EZWorksheet addedSheet = new EZWorksheet(worksheetPart.Worksheet, sheet, this);
            Worksheets.Add(addedSheet);

            nextAvailableSheetId++;

            return addedSheet;
        }

        public void Save()
        {
            spreadsheetDocument.WorkbookPart.Workbook.Save();

            spreadsheetDocument.Close();
        }

        public EZWorksheet GetSheet(string sheetName)
        {
            return Worksheets.Where(x => x.Sheet.Name == sheetName).FirstOrDefault();
        }

    }
}

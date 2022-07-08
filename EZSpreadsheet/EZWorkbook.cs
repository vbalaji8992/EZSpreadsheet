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
        internal SpreadsheetDocument SpreadsheetDocument { get; }
        internal Sheets Sheets { get; }
        internal List<EZWorksheet> Worksheets { get; }
        internal EZSharedString SharedString { get; }
        internal EZStylesheet StyleSheet { get; }
        internal uint NextAvailableSheetId { get; private set; } = 1;

        public EZWorkbook(string filepath)
        {
            SpreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            SpreadsheetDocument.AddWorkbookPart();
            SpreadsheetDocument.WorkbookPart!.Workbook = new Workbook();

            Sheets = SpreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            Worksheets = new List<EZWorksheet>();
            //AddSheet();

            SharedStringTablePart sharedStringPart = SpreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            SharedString = new EZSharedString(this, sharedStringPart);

            WorkbookStylesPart workbookStylesPart = SpreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            StyleSheet = new EZStylesheet(this, workbookStylesPart);
        }

        public EZWorksheet AddSheet(string? sheetName = null)
        {
            if (sheetName != null && GetSheet(sheetName) != null)
            {
                throw new Exception("Sheet already exists!");
            }

            EZWorksheet addedSheet = new EZWorksheet(this, sheetName);
            Worksheets.Add(addedSheet);

            NextAvailableSheetId++;

            return addedSheet;
        }

        public void Save()
        {
            SpreadsheetDocument.WorkbookPart?.Workbook.Save();

            SpreadsheetDocument.Close();
        }

        public EZWorksheet? GetSheet(string sheetName)
        {
            return Worksheets.Where(x => x.Sheet.Name == sheetName).FirstOrDefault();
        }

        public int GetSheetCount()
        {
            return Worksheets.Count();
        }
    }
}

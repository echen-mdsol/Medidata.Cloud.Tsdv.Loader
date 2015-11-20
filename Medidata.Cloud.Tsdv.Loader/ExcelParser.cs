using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.Helpers;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class ExcelParser : IExcelParser
    {
        private SpreadsheetDocument _doc;

        public IEnumerable<T> GetObjects<T>(string sheetName, bool hasHeaderRow = true) where T : class
        {
            var worksheet = FindWorksheet(_doc, sheetName);
            var worksheetParser = new SheetParser<T> {HasHeaderRow = hasHeaderRow};
            worksheetParser.Load(worksheet);

            return worksheetParser.GetObjects();
        }

        private Worksheet FindWorksheet(SpreadsheetDocument doc, string sheetName)
        {
            var attribute = SpreadsheetAttributeHelper.CreateSheetNameAttribute(sheetName);

            var sheet = doc.WorkbookPart.Workbook
                .Descendants<Sheet>()
                .Where(x => x.HasAttributes)
                .First(x => x.GetAttributes().Contains(attribute));
            var id = sheet.Id;
            var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(id);
            return worksheetPart.Worksheet;
        }

        public void Load(Stream stream)
        {
            _doc = SpreadsheetDocument.Open(stream, false);
        }

        public void Dispose()
        {
            if (_doc == null) return;
            _doc.Dispose();
            _doc = null;
        }
    }
}
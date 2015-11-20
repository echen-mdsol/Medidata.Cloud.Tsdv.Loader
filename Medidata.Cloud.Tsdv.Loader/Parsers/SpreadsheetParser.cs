using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.Helpers;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public class SpreadsheetParser : ISpreadsheetParser
    {
        private SpreadsheetDocument _doc;
        private bool _hasHeaderRow;

        public IEnumerable<T> RetrieveObjectsFromSheet<T>(string sheetName) where T : class
        {
            var attribute = SpreadsheetAttributeHelper.CreateSheetNameAttribute(sheetName);

            var sheet = _doc.WorkbookPart.Workbook
                .Descendants<Sheet>()
                .Where(x => x.HasAttributes)
                .First(x => x.GetAttributes().Contains(attribute));
            var id = sheet.Id;
            var worksheetPart = (WorksheetPart) _doc.WorkbookPart.GetPartById(id);
            var worksheet = worksheetPart.Worksheet;

            var worksheetParser = new WorksheetParser<T>();
            worksheetParser.Load(worksheet, _hasHeaderRow);

            return worksheetParser.GetObjects();
        }

        public void Load(Stream stream, bool hasHeaderRow = true)
        {
            _doc = SpreadsheetDocument.Open(stream, false);
            _hasHeaderRow = hasHeaderRow;
        }

        public void Dispose()
        {
            if (_doc == null) return;
            _doc.Dispose();
            _doc = null;
        }
    }
}
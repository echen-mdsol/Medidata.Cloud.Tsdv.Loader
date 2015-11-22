using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    public class ExcelParser : IExcelParser
    {
        private readonly ICellTypeValueConverterFactory _converterFactory;
        private SpreadsheetDocument _doc;

        public ExcelParser(ICellTypeValueConverterFactory converterFactory)
        {
            if (converterFactory == null) throw new ArgumentNullException("converterFactory");
            _converterFactory = converterFactory;
        }

        public IEnumerable<T> GetObjects<T>(string sheetName, bool hasHeaderRow = true) where T : class
        {
            var worksheet = FindWorksheet(_doc, sheetName);
            var worksheetParser = new SheetParser<T>(_converterFactory) {HasHeaderRow = hasHeaderRow};
            worksheetParser.Load(worksheet);

            return worksheetParser.GetObjects();
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

        private Worksheet FindWorksheet(SpreadsheetDocument doc, string sheetName)
        {
            var attribute = SpreadsheetAttributeHelper.CreateSheetNameAttribute(sheetName);

            var sheet = doc.WorkbookPart.Workbook
                .Descendants<Sheet>()
                .Where(x => x.HasAttributes)
                .First(x => x.GetAttributes().Contains(attribute));
            var id = sheet.Id;
            var worksheetPart = (WorksheetPart) doc.WorkbookPart.GetPartById(id);
            return worksheetPart.Worksheet;
        }
    }
}
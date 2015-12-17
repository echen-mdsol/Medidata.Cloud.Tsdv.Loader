using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Medidata.Cloud.ExcelLoader.Helpers;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public class ExcelParser : IExcelParser
    {
        private SpreadsheetDocument _doc;

        public IEnumerable<ExpandoObject> GetObjects(ISheetDefinition sheetDefinition, ISheetParser sheetParser)
        {
            var worksheet = _doc.GetWorksheetByName(sheetDefinition.Name);
            return sheetParser.GetObjects(worksheet, sheetDefinition);
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
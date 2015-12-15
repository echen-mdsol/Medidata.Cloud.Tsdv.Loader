using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ImpromptuInterface;
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

        public IEnumerable<ExpandoObject> GetObjects(ISheetDefinition sheetDefinition)
        {
            var worksheet = _doc.GetWorksheetByName(sheetDefinition.Name);
            var worksheetParser = new SheetParser(_converterFactory);
            worksheetParser.Load(worksheet);

            return worksheetParser.GetObjects(sheetDefinition);
        }

        public IEnumerable<T> GetObjects<T>(ISheetDefinition sheetDefinition) where T : class
        {
            return GetObjects(sheetDefinition).Select(x => x.ActLike<T>());
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
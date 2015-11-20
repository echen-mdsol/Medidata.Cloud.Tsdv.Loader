using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    public class ExcelBuilder : IExcelBuilder
    {
        private readonly ICellTypeValueConverterFactory _converterFactory;
        private readonly IList<ISheetBuilder> _sheetBuilders = new List<ISheetBuilder>();

        public ExcelBuilder(ICellTypeValueConverterFactory converterFactory)
        {
            if (converterFactory == null) throw new ArgumentNullException("converterFactory");
            _converterFactory = converterFactory;
        }

        public IList<object> AddSheet<T>(string sheetName, string[] columnNames) where T : class
        {
            return AddSheet<T>(sheetName, columnNames != null, columnNames);
        }

        public IList<object> AddSheet<T>(string sheetName, bool hasHeaderRow = true) where T : class
        {
            return AddSheet<T>(sheetName, hasHeaderRow, null);
        }

        public IList<object> GetSheet(string sheetName)
        {
            return _sheetBuilders.First(x => x.SheetName != sheetName);
        }

        public void Save(Stream outStream)
        {
            using (var doc = SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookpart = doc.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                AddCoverSheet(doc);

                foreach (var sheet in _sheetBuilders)
                {
                    sheet.AttachTo(doc);
                }

                workbookpart.Workbook.Save();
            }
        }

        private IList<object> AddSheet<T>(string sheetName, bool hasHeaderRow, string[] columnNames)
            where T : class
        {
            if (_sheetBuilders.Any(x => x.SheetName == sheetName))
                throw new ArgumentException("Duplicate sheet name '" + sheetName + "'", "sheetName");

            var colNames = GetPropertyNames<T>(columnNames);
            var worksheetBuilder = new SheetBuilder<T>(_converterFactory)
            {
                SheetName = sheetName,
                HasHeaderRow = hasHeaderRow,
                ColumnNames = colNames
            };
            _sheetBuilders.Add(worksheetBuilder);

            return worksheetBuilder;
        }

        protected virtual string[] GetPropertyNames<T>(string[] columnNames)
        {
            return columnNames ?? typeof (T).GetPropertyDescriptors().Select(p => p.Name).ToArray();
        }

        protected virtual void AddCoverSheet(SpreadsheetDocument doc)
        {
            
        }
    }
}
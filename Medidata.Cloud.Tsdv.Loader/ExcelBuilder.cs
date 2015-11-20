using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.Helpers;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class ExcelBuilder : IExcelBuilder
    {
        private readonly IList<ISheetBuilder> _sheetBuilders = new List<ISheetBuilder>();

        public IList<object> AddSheet<T>(string sheetName, string[] columnNames) where T : class
        {
            return AddSheet<T>(sheetName, columnNames != null, columnNames);
        }

        public IList<object> AddSheet<T>(string sheetName, bool hasHeaderRow) where T : class
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

                var coverBuilder = new CoverSheetBuilder();
                coverBuilder.AttachTo(doc);

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

            var colNames = columnNames ?? GetPropertyNames<T>();
            var worksheetBuilder = new SheetBuilder<T>
            {
                SheetName = sheetName,
                HasHeaderRow = hasHeaderRow,
                ColumnNames = colNames
            };
            _sheetBuilders.Add(worksheetBuilder);

            return worksheetBuilder;
        }

        private string[] GetPropertyNames<T>()
        {
            return typeof (T).GetPropertyDescriptors().Select(p => p.Name).ToArray();
        }
    }
}
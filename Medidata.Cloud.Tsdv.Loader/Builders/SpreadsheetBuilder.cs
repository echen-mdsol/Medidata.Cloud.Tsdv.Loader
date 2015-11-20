using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.Extensions;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class SpreadsheetBuilder : ISpreadsheetBuilder
    {
        private readonly IDictionary<string, IWorksheetBuilder> _sheets =
            new Dictionary<string, IWorksheetBuilder>(StringComparer.OrdinalIgnoreCase);

        private bool _hasHeaderRow;

        public IList<object> EnsureWorksheet<T>(string sheetName, bool hasHeaderRow = true, string[] columnNames = null)
            where T : class
        {
            _hasHeaderRow = hasHeaderRow;
            IWorksheetBuilder worksheetBuilder;
            if (!_sheets.TryGetValue(sheetName, out worksheetBuilder))
            {
                var colNames = columnNames ?? GetPropertyNames<T>();
                worksheetBuilder = new WorksheetBuilder<T> {ColumnNames = colNames};
                _sheets.Add(sheetName, worksheetBuilder);
            }
            return worksheetBuilder;
        }

        public void Save(Stream outStream)
        {
            using (var doc = SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookpart = doc.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // TODO: Copy cover sheet from the resource
//            new CoverWorksheetBuilder().AppendWorksheet(doc, "Cover");

                foreach (var kvp in _sheets)
                {
                    var worksheetBuilder = kvp.Value;
                    worksheetBuilder.AppendWorksheet(doc, _hasHeaderRow, kvp.Key);
                }

                workbookpart.Workbook.Save();
            }
        }

        private string[] GetPropertyNames<T>()
        {
            return typeof (T).GetPropertyDescriptors().Select(p => p.Name).ToArray();
        }
    }
}
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

        public IList<T> EnsureWorksheet<T>(string sheetName, string[] columnNames = null) where T : class
        {
            IWorksheetBuilder worksheetBuilder;
            if (!_sheets.TryGetValue(sheetName, out worksheetBuilder))
            {
                var colNames = columnNames ?? GetPropertyNames<T>();
                worksheetBuilder = new WorksheetBuilder<T>() { ColumnNames = colNames };
                _sheets.Add(sheetName, worksheetBuilder);
            }
            return (IList<T>)worksheetBuilder;
        }

        private string[] GetPropertyNames<T>()
        {
            return typeof(T).GetPropertyDescriptors().Select(p => p.Name).ToArray();
        }

        public SpreadsheetDocument Save(Stream outStream)
        {
            var doc = SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook);
            var workbookpart = doc.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // TODO: Copy cover sheet from the resource
//            new CoverWorksheetBuilder().AppendWorksheet(doc, "Cover");

            foreach (var kvp in _sheets)
            {
                var worksheetBuilder = kvp.Value;
                worksheetBuilder.AppendWorksheet(doc, kvp.Key);
            }

            workbookpart.Workbook.Save();

            doc.Close();
            return doc;
        }
    }
}
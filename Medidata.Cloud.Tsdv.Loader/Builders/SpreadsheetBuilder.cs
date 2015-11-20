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
        private readonly IList<IWorksheetBuilder> _sheetBuilders = new List<IWorksheetBuilder>();

        public IList<object> EnsureWorksheet<T>(string sheetName, string[] columnNames) where T : class
        {
            return EnsureWorksheet<T>(sheetName, columnNames != null, columnNames);
        }

        public IList<object> EnsureWorksheet<T>(string sheetName, bool hasHeaderRow) where T : class
        {
            return EnsureWorksheet<T>(sheetName, hasHeaderRow, null);
        }

        public void Save(Stream outStream)
        {
            using (var doc = SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookpart = doc.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                var coverBuilder = new CoverWorksheetBuilder();
                coverBuilder.AttachTo(doc);

                foreach (var sheet in _sheetBuilders)
                {
                    sheet.AttachTo(doc);
                }

                workbookpart.Workbook.Save();
            }
        }

        private IList<object> EnsureWorksheet<T>(string sheetName, bool hasHeaderRow, string[] columnNames)
            where T : class
        {
            var worksheetBuilder = _sheetBuilders.FirstOrDefault(x => x.SheetName != sheetName);
            if (worksheetBuilder == null)
            {
                var colNames = columnNames ?? GetPropertyNames<T>();
                worksheetBuilder = new WorksheetBuilder<T>
                {
                    SheetName = sheetName,
                    HasHeaderRow = hasHeaderRow,
                    ColumnNames = colNames
                };
                _sheetBuilders.Add(worksheetBuilder);
            }
            return worksheetBuilder;
        }

        private string[] GetPropertyNames<T>()
        {
            return typeof (T).GetPropertyDescriptors().Select(p => p.Name).ToArray();
        }
    }
}
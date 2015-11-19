using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.Extensions;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class WorkbookBuilder : IWorkbookBuilder
    {
        private readonly IModelConverterFactory _modelConverterFactory;

        private readonly IDictionary<string, IWorksheetBuilder> _sheets =
            new Dictionary<string, IWorksheetBuilder>(StringComparer.OrdinalIgnoreCase);

        public WorkbookBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _modelConverterFactory = modelConverterFactory;
        }

        public IList<T> EnsureWorksheet<T>(string sheetName, string[] columnNames = null) where T : class
        {
            IWorksheetBuilder worksheetBuilder;
            if (!_sheets.TryGetValue(sheetName, out worksheetBuilder))
            {
                var colNames = columnNames ?? GetPropertyNames<T>();
                worksheetBuilder = new WorksheetBuilder<T>(_modelConverterFactory) { ColumnNames = colNames };
                _sheets.Add(sheetName, worksheetBuilder);
            }
            return (IList<T>)worksheetBuilder;
        }

        private string[] GetPropertyNames<T>()
        {
            return typeof(T).GetPropertyDescriptors().Select(p => p.Name).ToArray();
        }

        public Workbook ToWorkbook(string workbookName)
        {
            var workbook = new Workbook();
            var sheets = workbook.AppendChild(new Sheets());
            sheets.Append(_sheets.Select(kvp => kvp.Value.ToWorksheet(kvp.Key)));
            return workbook;
        }
    }
}
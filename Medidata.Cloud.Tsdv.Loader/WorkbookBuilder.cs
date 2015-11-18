using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class WorkbookBuilder : IWorkbookBuilder
    {
        private readonly IModelConverterFactory _modelConverterFactory;

        private readonly IDictionary<string, object> _sheets =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        public WorkbookBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _modelConverterFactory = modelConverterFactory;
        }

        public IWorksheetBuilder<T> EnsureWorksheet<T>(string sheetName) where T : class
        {
            object worksheetBuilder;
            if (!_sheets.TryGetValue(sheetName, out worksheetBuilder))
            {
                worksheetBuilder = new WorksheetBuilder<T>(_modelConverterFactory);
                _sheets.Add(sheetName, worksheetBuilder);
            }
            return (IWorksheetBuilder<T>)worksheetBuilder;
        }

        public Workbook ToWorkbook(string workbookName)
        {
            throw new NotImplementedException();
        }
    }
}
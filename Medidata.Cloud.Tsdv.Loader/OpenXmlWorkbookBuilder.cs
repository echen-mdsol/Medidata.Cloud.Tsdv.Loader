using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class OpenXmlWorkbookBuilder : IOpenXmlWorkbookBuilder
    {
        private readonly IModelConverterFactory _modelConverterFactory;

        private readonly IDictionary<string, object> _sheets =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        public OpenXmlWorkbookBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _modelConverterFactory = modelConverterFactory;
        }

        public IOpenXmlWorksheetBuilder<T> EnsureWorksheet<T>(string sheetName) where T : class
        {
            object worksheetBuilder;
            if (!_sheets.TryGetValue(sheetName, out worksheetBuilder))
            {
                worksheetBuilder = new OpenXmlWorksheetBuilder<T>(_modelConverterFactory);
                _sheets.Add(sheetName, worksheetBuilder);
            }
            return (IOpenXmlWorksheetBuilder<T>)worksheetBuilder;
        }

        public Workbook ToOpenXmlWorkbook(string workbookName)
        {
            throw new NotImplementedException();
        }
    }
}
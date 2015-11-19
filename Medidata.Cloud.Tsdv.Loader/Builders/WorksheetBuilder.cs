using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class WorksheetBuilder<T> : List<T>, IWorksheetBuilder
    {
        private readonly IModelConverterFactory _modelConverterFactory;
        private readonly Type _objectType;

        public WorksheetBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _objectType = typeof(T);
            _modelConverterFactory = modelConverterFactory;
        }

        public string[] ColumnNames { get; set; }

        public Sheet ToWorksheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException("sheetName cannot be empty");
            var converter = _modelConverterFactory.ProduceConverter(_objectType);
            var models = this.Select(x => converter.ConvertToModel(x));

            // ExcelHelper
            var sheet = new Sheet { Name = sheetName };
            return sheet;
        }
    }
}
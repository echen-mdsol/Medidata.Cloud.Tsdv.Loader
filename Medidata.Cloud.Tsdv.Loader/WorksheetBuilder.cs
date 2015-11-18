using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
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

        public Sheet ToWorksheet(string name)
        {
            var converter = _modelConverterFactory.ProduceConverter(_objectType);
            var models = this.Select(x => converter.ConvertToModel(x));
            
            // ExcelHelper
            var sheet = new Sheet {Name = name};
            return sheet;
        }
    }
}
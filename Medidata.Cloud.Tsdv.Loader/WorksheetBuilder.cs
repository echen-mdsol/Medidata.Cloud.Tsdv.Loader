using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class WorksheetBuilder<T> : IWorksheetBuilder<T> where T : class
    {
        private readonly IModelConverterFactory _modelConverterFactory;
        private readonly IList<T> _objects = new List<T>();
        private Type _objectType;

        public WorksheetBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _objectType = typeof(T);
            _modelConverterFactory = modelConverterFactory;
        }


        public void AddObject(T target)
        {
            _objects.Add(target);
        }

        public Worksheet ToWorksheet()
        {
            var converter = _modelConverterFactory.ProduceConverter(_objectType);
            var models = _objects.Select(x => converter.ConvertToModel(x));

            // ExcelHelper
            throw new NotImplementedException();
        }
    }
}
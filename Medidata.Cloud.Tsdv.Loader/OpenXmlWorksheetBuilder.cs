using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class OpenXmlWorksheetBuilder<T> : IOpenXmlWorksheetBuilder<T> where T : class
    {
        private readonly IModelConverterFactory _modelConverterFactory;
        private readonly IList<T> _objects = new List<T>();
        private Type _objectType;

        public OpenXmlWorksheetBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _objectType = typeof(T);
            _modelConverterFactory = modelConverterFactory;
        }


        public void AddObject(T target)
        {
            _objects.Add(target);
        }

        public Worksheet ToOpenXmlWorksheet()
        {
            var converter = _modelConverterFactory.ProduceConverter(_objectType);
            var models = _objects.Select(x => converter.ConvertToModel(x));
            throw new NotImplementedException();
        }
    }
}
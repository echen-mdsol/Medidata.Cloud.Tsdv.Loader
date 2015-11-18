using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class OpenXmlWorksheetParser<T> : IOpenXmlWorksheetParser<T> where T : class
    {
        private readonly IModelConverterFactory _modelConverterFactory;
        private readonly IList<object> _objects = new List<object>();

        public OpenXmlWorksheetParser(Type modelType, IModelConverterFactory modelConverterFactory)
        {
            if (modelType == null) throw new ArgumentNullException("modelType");
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _modelConverterFactory = modelConverterFactory;
        }

        public IEnumerable<T> GetObjects()
        {
            throw new NotImplementedException();
        }

        public void Load(Worksheet worksheet)
        {
            throw new NotImplementedException();
        }
    }
}
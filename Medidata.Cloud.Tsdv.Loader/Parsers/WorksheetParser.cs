using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public class WorksheetParser<T> : IWorksheetParser<T> where T : class
    {
        private readonly IModelConverterFactory _modelConverterFactory;
        private readonly IList<object> _objects = new List<object>();

        public WorksheetParser(Type modelType, IModelConverterFactory modelConverterFactory)
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
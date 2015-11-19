using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public class WorksheetParser<T> : IWorksheetParser<T> where T : class
    {
        private readonly IList<object> _objects = new List<object>();

        public WorksheetParser(Type modelType)
        {
            if (modelType == null) throw new ArgumentNullException("modelType");
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
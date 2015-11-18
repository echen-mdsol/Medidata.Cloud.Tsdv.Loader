using System;
using System.Collections.Generic;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class WorkbookModelBuilder : IWorkbookModelBuilder
    {
        private readonly IModelConverterFactory _modelConverterFactory;

        private readonly IDictionary<string, IWorksheetModelBuilder> _sheets =
            new Dictionary<string, IWorksheetModelBuilder>(StringComparer.OrdinalIgnoreCase);

        public WorkbookModelBuilder(IModelConverterFactory modelConverterFactory)
        {
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            _modelConverterFactory = modelConverterFactory;
        }

        public void AddWorksheet<T>(string sheetName) where T : class
        {
            _sheets.Add(sheetName, new WorksheetModelBuilder(typeof (T), _modelConverterFactory));
        }

        public IWorksheetModelBuilder GetWorksheet<TInterface>(string sheetName) where TInterface : class
        {
            return _sheets[sheetName];
        }

        public void RemoveWorksheet(string sheetName)
        {
            _sheets.Remove(sheetName);
        }

        public bool ContainsWorksheet(string sheetName)
        {
            return _sheets.ContainsKey(sheetName);
        }

        public object ToModel()
        {
            throw new NotImplementedException();
        }
    }
}
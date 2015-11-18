using System;
using System.Collections.Generic;
using System.Linq;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class WorksheetModelBuilder : IWorksheetModelBuilder
    {
        private readonly IModelConverterFactory _modelConverterFactory;
        private readonly IList<object> _objects = new List<object>();

        public WorksheetModelBuilder(Type interfaceType, IModelConverterFactory modelConverterFactory)
        {
            if (interfaceType == null) throw new ArgumentNullException("interfaceType");
            if (modelConverterFactory == null) throw new ArgumentNullException("modelConverterFactory");
            InterfaceType = interfaceType;
            _modelConverterFactory = modelConverterFactory;
        }

        public Type InterfaceType { get; private set; }

        public void AddObject<TInterface>(TInterface target) where TInterface : class
        {
            if (typeof (TInterface) != InterfaceType)
                throw new NotSupportedException("Object must be an instance of " + InterfaceType);
            _objects.Add(target);
        }

        public IEnumerable<TModel> GetModels<TModel>() where TModel : class
        {
            var converter = _modelConverterFactory.ProduceConverter(InterfaceType);
            if (converter.ViewModelType != typeof (TModel))
                throw new NotSupportedException(typeof (TModel) + " isn't a defined model type.");
            return _objects.Select(x => converter.ConvertToViewModel(x)).Cast<TModel>();
        }
    }
}
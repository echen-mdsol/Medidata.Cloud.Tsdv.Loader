using System;
using System.Collections.Generic;
using System.Linq;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class ModelConverterFactory : IModelConverterFactory
    {
        private readonly IDictionary<Type, IModelConverter> _converters = new Dictionary<Type, IModelConverter>();

        public ModelConverterFactory(IModelConverter[] customConverters)
        {
            if (customConverters == null) return;
            _converters = _converters
                .Union(customConverters.Select(x => new KeyValuePair<Type, IModelConverter>(x.InterfaceType, x)))
                .ToDictionary(x => x.Key, x => x.Value);
        }

        public IModelConverter ProduceConverter(Type type)
        {
            return _converters[type];
        }
    }
}
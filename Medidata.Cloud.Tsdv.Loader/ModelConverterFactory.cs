using System;
using System.Collections.Generic;
using System.Linq;
using Medidata.Cloud.Tsdv.Loader.ModelConverters;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class ModelConverterFactory : IModelConverterFactory
    {
        private readonly IDictionary<Type, IModelConverter> _converters = new Dictionary<Type, IModelConverter>();

        public ModelConverterFactory(IModelConverter[] customConverters)
        {
            var converters = new IModelConverter[]
            {
                new BlockPlanConverter(),
                new TierFormConverter()
            };
            if (customConverters == null) return;
            _converters = converters.Union(customConverters).ToDictionary(x => x.InterfaceType, x => x);
        }

        public IModelConverter ProduceConverter(Type type)
        {
            return _converters[type];
        }
    }
}
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
            IEnumerable<IModelConverter> converters = new IModelConverter[]
            {
                new BlockPlanConverter(),
                new BlockPlanSettingConverter(), 
                new CustomTierConverter(), 
                new TierFormConverter(),
                new ExcludedStatusConverter(), 
                new RuleConverter()
            };
            if(customConverters!= null)
            {
                converters = converters.Union(customConverters);
            }
            _converters = converters.ToDictionary(x => x.InterfaceType, x => x);
        }

        public IModelConverter ProduceConverter(Type type)
        {
            return _converters[type];
        }
    }
}
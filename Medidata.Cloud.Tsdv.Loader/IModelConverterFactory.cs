using System;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IModelConverterFactory
    {
        IModelConverter ProduceConverter(Type type);
    }
}
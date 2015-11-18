using System;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IModelConverter
    {
        Type InterfaceType { get; }
        Type ModelType { get; }
        object ConvertToModel(object target);
        object ConvertFromModel(object model);
    }
}
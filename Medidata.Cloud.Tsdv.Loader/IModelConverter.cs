using System;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IModelConverter
    {
        Type InterfaceType { get; }
        Type ViewModelType { get; }
        object ConvertToViewModel(object target);
        object ConvertFromViewModel(object viewModel);
    }
}
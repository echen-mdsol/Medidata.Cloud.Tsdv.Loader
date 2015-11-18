using System;
using System.Collections.Generic;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IWorksheetModelBuilder
    {
        Type InterfaceType { get; }
        void AddObject<TInterface>(TInterface target) where TInterface : class;
        IEnumerable<TModel> GetModels<TModel>() where TModel : class;
    }
}
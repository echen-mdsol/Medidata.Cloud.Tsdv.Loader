using System.Collections.Generic;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetInfo<T> where T : SheetModel
    {
        ISheetDefinition Definition { get; }
        IList<T> Data { get; }
    }
}
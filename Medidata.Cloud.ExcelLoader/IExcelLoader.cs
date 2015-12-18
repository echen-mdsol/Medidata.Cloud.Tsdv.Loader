using System.Collections.Generic;
using System.IO;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelLoader
    {
        void Save(Stream outStream);
        void Load(Stream source);
        ISheetDefinition SheetDefinition<T>() where T : SheetModel;
        IList<T> SheetData<T>() where T : SheetModel;
    }
}
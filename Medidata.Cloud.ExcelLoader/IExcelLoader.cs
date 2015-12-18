using System.IO;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelLoader
    {
        void Save(Stream outStream);
        void Load(Stream source);
        ISheetInfo<T> Sheet<T>() where T : SheetModel;
    }
}
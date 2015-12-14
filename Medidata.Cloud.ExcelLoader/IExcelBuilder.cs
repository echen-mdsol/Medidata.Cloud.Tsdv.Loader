using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelBuilder
    {
        IList<object> DefineSheet(ISheetDefinition sheetDefinition, ISheetBuilder sheetBuilder);
        void Save(Stream outStream);
    }
}
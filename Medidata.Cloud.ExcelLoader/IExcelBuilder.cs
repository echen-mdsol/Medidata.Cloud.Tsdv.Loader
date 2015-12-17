using System.Collections.Generic;
using System.IO;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelBuilder
    {
        void AddSheet(ISheetDefinition sheetDefinition, IEnumerable<SheetModel> models, ISheetBuilder sheetBuilder);
        void Save(Stream outStream);
    }
}
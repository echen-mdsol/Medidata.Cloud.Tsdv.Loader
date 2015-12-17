using System.Collections.Generic;
using System.IO;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelBuilder
    {
        IList<SheetDefinitionModelBase> DefineSheet(ISheetDefinition sheetDefinition, ISheetBuilder sheetBuilder);
        void Save(Stream outStream);
    }
}
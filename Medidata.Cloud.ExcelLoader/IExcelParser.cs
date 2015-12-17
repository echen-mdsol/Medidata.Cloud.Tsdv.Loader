using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelParser : IDisposable
    {
        IEnumerable<ExpandoObject> GetObjects(ISheetDefinition sheetDefinition);
        IEnumerable<T> GetObjects<T>(ISheetDefinition sheetDefinition) where T : SheetModel;
        void Load(Stream stream);
    }
}
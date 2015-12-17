using System.Collections.Generic;
using System.Dynamic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    internal interface ISheetParser
    {
        IEnumerable<ExpandoObject> GetObjects(Worksheet worksheet, ISheetDefinition sheetDefinition);
    }
}
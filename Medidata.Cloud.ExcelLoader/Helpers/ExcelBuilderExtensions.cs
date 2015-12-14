using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class ExcelBuilderExtensions
    {
        public static IList<object> DefineSheet<T>(this IExcelBuilder target, string sheetName, ISheetBuilder sheetBuilder)
        {
            var sheetDefinition = typeof (T).ToSheetDefinition(sheetName);
            return target.DefineSheet(sheetDefinition, sheetBuilder);
        }
    }
}
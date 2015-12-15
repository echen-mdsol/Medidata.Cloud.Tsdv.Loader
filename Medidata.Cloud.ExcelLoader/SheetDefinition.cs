using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public class SheetDefinition: ISheetDefinition
    {
        public string Name { get; set; }
        public IEnumerable<IColumnDefinition> ColumnDefinitions { get; set; }
    }
}
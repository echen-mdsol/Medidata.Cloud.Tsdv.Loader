using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public class SheetDefinition : ISheetDefinition
    {
        public string Name { get; set; }
        public bool AcceptExtraProperties { get; set; }
        public IList<IColumnDefinition> ColumnDefinitions { get; set; }
    }
}
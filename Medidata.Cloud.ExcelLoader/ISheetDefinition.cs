using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetDefinition
    {
        string Name { get; }
        IEnumerable<string> Headers { get; }
        bool AutoFilter { get; }
        IEnumerable<IColumnDefinition> ColumnDefinitions { get; }
    }

    public class SheetDefinition: ISheetDefinition
    {
        public string Name { get; set; }
        public IEnumerable<string> Headers { get; set; }
        public bool AutoFilter { get; set; }
        public IEnumerable<IColumnDefinition> ColumnDefinitions { get; set; }
    }
}
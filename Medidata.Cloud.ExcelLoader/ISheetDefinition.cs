using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetDefinition
    {
        string Name { get; }
        IEnumerable<IColumnDefinition> ColumnDefinitions { get; }
    }
}
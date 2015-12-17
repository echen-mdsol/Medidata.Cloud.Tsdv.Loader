using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetDefinition
    {
        string Name { get; }
        bool AcceptExtraProperties { get; }
        IEnumerable<IColumnDefinition> ColumnDefinitions { get; }
        ISheetDefinition AddColumn(IColumnDefinition columnDefinition);
    }
}
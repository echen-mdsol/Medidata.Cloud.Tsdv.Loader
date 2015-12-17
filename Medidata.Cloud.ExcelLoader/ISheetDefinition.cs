using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetDefinition
    {
        string Name { get; }
        bool AcceptExtraProperties { get; }
        IEnumerable<IColumnDefinition> ColumnDefinitions { get; }
        ISheetDefinition DefineColumns<T>(params string[] headers);
        ISheetDefinition AddColumn(IColumnDefinition columnDefinition);
    }
}
using System.Collections.Generic;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetDefinition
    {
        string Name { get; }
        bool AcceptExtraProperties { get; }
        IList<IColumnDefinition> ColumnDefinitions { get; }
    }
}
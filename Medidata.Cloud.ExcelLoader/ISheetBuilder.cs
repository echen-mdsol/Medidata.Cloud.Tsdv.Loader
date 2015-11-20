using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder : IList<object>
    {
        string SheetName { get; }
        bool HasHeaderRow { get; }
        string[] ColumnNames { get; }
        void AttachTo(SpreadsheetDocument doc);
    }
}
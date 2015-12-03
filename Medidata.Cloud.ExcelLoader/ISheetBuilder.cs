using System.Collections;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder : IList
    {
        string SheetName { get; }
        bool HasHeaderRow { get; }
        string[] ColumnNames { get; }
        void AttachTo(SpreadsheetDocument doc);
        string HeaderStyleName { get; set; }
        string TextStyleName { get; set; }
    }
}
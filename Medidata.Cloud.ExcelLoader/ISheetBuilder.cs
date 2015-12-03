using System.Collections;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder : IList
    {
        string SheetName { get; set; }
        bool HasHeaderRow { get; set; }
        string[] ColumnNames { get; set; }
        void AttachTo(SpreadsheetDocument doc);
    }
}
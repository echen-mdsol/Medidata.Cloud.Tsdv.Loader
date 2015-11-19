using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface IWorksheetBuilder
    {
        string[] ColumnNames { get; set; }
        void AppendWorksheet(SpreadsheetDocument doc, string sheetName);
    }
}
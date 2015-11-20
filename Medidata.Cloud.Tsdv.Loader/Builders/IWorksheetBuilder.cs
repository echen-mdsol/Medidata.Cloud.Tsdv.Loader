using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface IWorksheetBuilder
    {
        string[] ColumnNames { get; set; }
        void AppendWorksheet(SpreadsheetDocument doc, bool hasHeaderRow, string sheetName);
    }
}
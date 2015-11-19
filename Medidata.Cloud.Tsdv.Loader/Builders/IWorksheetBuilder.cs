using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface IWorksheetBuilder
    {
        string[] ColumnNames { get; set; }
        Sheet ToWorksheet(string sheetName);
    }
}
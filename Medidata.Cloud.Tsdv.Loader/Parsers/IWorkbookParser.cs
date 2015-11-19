using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public interface IWorkbookParser
    {
        IWorksheetParser<T> GetWorksheet<T>(string sheetName) where T : class;
        void Load(Workbook workbook);
    }
}
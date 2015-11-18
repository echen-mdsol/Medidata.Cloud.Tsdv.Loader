using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IOpenXmlWorkbookBuilder
    {
        IOpenXmlWorksheetBuilder<T> EnsureWorksheet<T>(string sheetName) where T : class;
        Workbook ToOpenXmlWorkbook(string workbookName);
    }
}
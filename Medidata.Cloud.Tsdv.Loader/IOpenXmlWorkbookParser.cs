using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IOpenXmlWorkbookParser
    {
        IOpenXmlWorksheetParser<T> GetWorksheet<T>(string sheetName) where T : class;
        void Load(Workbook workbook);
    }
}
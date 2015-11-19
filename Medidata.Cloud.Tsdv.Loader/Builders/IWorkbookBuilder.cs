using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface IWorkbookBuilder
    {
        IList<T> EnsureWorksheet<T>(string sheetName, string [] columnNames = null) where T : class;
        Workbook ToWorkbook(string workbookName);
    }
}
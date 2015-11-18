using System;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class WorkbookParser : IWorkbookParser
    {
        public IWorksheetParser<T> GetWorksheet<T>(string sheetName) where T : class
        {
            throw new NotImplementedException();
        }

        public void Load(Workbook workbook)
        {
            
        }
    }
}
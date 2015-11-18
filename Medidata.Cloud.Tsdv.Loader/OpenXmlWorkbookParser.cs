using System;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public class OpenXmlWorkbookParser : IOpenXmlWorkbookParser
    {
        public IOpenXmlWorksheetParser<T> GetWorksheet<T>(string sheetName) where T : class
        {
            throw new NotImplementedException();
        }

        public void Load(Workbook workbook)
        {
            
        }
    }
}
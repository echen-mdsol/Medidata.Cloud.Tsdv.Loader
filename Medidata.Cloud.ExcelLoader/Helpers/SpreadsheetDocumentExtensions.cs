using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class SpreadsheetDocumentExtensions
    {
        public static Worksheet GetWorksheetByName(this SpreadsheetDocument doc, string sheetName)
        {
            var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>().First(x => x.Name == sheetName);
            var sheetPartId = sheet.Id;
            var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheetPartId);
            return worksheetPart.Worksheet;
        }

        public static SheetData GetSheetDataByName(this SpreadsheetDocument doc, string sheetName)
        {
            var worksheet = GetWorksheetByName(doc, sheetName);
            return worksheet.GetFirstChild<SheetData>();
        }

        public static uint GetStyleIndex(this SpreadsheetDocument doc, string styleName)
        {
            var provider = new CellStyleProvider();
            return provider.GetStyleIndex(doc, styleName);
        }
    }
}
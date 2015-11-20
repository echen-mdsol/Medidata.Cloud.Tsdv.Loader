using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public static class WorksheetBuilderHelper
    {
        public static void AppendWorksheet(SpreadsheetDocument doc, Worksheet worksheet, string sheetName)
        {
            var sheets = doc.WorkbookPart.Workbook.Sheets ?? doc.WorkbookPart.Workbook.AppendChild(new Sheets());

            var worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = worksheet;

            var sheetId = (uint) (1 + sheets.Count());
            var sheet = new Sheet
            {
                Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = sheetName
            };
            sheets.Append(sheet);
        }
    }
}
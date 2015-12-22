using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader.SheetDecorators
{
    public class AutoFitColumnSheetDecorator : ISheetBuilderDecorator
    {
        public ISheetBuilder Decorate(ISheetBuilder target)
        {
            var originalFunc = target.BuildSheet;
            target.BuildSheet = (models, sheetDefinition, doc) =>
            {
                originalFunc(models, sheetDefinition, doc);

                AddColumns(doc, sheetDefinition.Name);
            };
            return target;
        }

        private void AddColumns(SpreadsheetDocument doc, string sheetName)
        {
            var sheetData = doc.GetSheetDataByName(sheetName);
            var headerRow = sheetData.Descendants<Row>().FirstOrDefault();
            if (headerRow == null) return;

            var columns = new Columns();
            columns.Append(headerRow.Descendants<Cell>().Select((x, i) => CreateColumn(doc, x, i)));

            var sheet = doc.GetWorksheetByName(sheetName);
            sheet.InsertAt(columns, 0);
        }

        private Column CreateColumn(SpreadsheetDocument doc, Cell cell, int cellIndex)
        {
            var column = new Column
                         {
                             Min = (uint) (cellIndex + 1),
                             Max = (uint) (cellIndex + 1),
                             Width = 20D, // TODO: Needs a way that can automaticaly calculate column width based on the column contents
                             CustomWidth = true
                         };
            return column;
        }
    }
}
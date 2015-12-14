using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader.SheetDecorators
{
    public class AutoFilterSheetDecorator : ISheetBuilderDecorator
    {
        public ISheetBuilder Decorate(ISheetBuilder target)
        {
            var originalFunc = target.BuildSheet;
            target.BuildSheet = (models, sheetDefinition, doc) =>
            {
                originalFunc(models, sheetDefinition, doc);
                var sheetName = sheetDefinition.Name;
                var sheetData = doc.GetSheetDataByName(sheetName);

                var firstRow = sheetData.Descendants<Row>().FirstOrDefault();
                var numberOfColumns = firstRow == null ? 0 : firstRow.Descendants<Cell>().Count();

                //No data, no filter will be created
                if (numberOfColumns == 0) return;

                var filter = new AutoFilter
                {
                    Reference = string.Format("{0}1:{1}1", GetColumnName(1), GetColumnName(numberOfColumns))
                };

                var worksheet = doc.GetWorksheetByName(sheetName);
                worksheet.AppendChild(filter);
            };

            return target;
        }


        //Generate excel style of cell index number, e.g. A5, B6, AC123, BD5448
        private string GetColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
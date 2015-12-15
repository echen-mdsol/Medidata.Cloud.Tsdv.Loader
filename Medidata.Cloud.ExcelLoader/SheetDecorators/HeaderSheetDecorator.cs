using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader.SheetDecorators
{
    public class HeaderSheetDecorator : ISheetBuilderDecorator
    {
        public ISheetBuilder Decorate(ISheetBuilder target)
        {
            var originalFunc = target.BuildSheet;
            target.BuildSheet = (models, sheetDefinition, doc) =>
            {
                originalFunc(models, sheetDefinition, doc);
                var headers = sheetDefinition.ColumnDefinitions.Select(x => x.HeaderName ?? x.PropertyName);
                var headerRow = CreateHeaderRow(headers);

                var sheetData = doc.GetSheetDataByName(sheetDefinition.Name);
                sheetData.InsertAt(headerRow, 0);
            };
            return target;
        }

        protected virtual Row CreateHeaderRow(IEnumerable<string> headers)
        {
            var row = new Row();
            foreach (var columnName in headers)
            {
                var cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(columnName),
                };
                row.AppendChild(cell);
            }
            return row;
        }
    }
}
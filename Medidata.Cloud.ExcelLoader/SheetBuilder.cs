using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.Helpers;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public class SheetBuilder : ISheetBuilder
    {
        private readonly ICellTypeValueConverterManager _converterManager;

        public SheetBuilder(ICellTypeValueConverterManager converterManager, params ISheetBuilderDecorator[] decorators)
        {
            if (converterManager == null) throw new ArgumentNullException("converterManager");
            _converterManager = converterManager;
            BuildSheet = BuildSheetFunc;
            BuildRow = BuildRowFromExpandoObject;
            foreach (var decorator in decorators)
            {
                decorator.Decorate(this);
            }
        }

        public Action<IEnumerable<SheetModel>, ISheetDefinition, SpreadsheetDocument> BuildSheet { get; set; }

        public Func<SheetModel, ISheetDefinition, Row> BuildRow { get; set; }

        private Row BuildRowFromExpandoObject(SheetModel model, ISheetDefinition sheetDefinition)
        {
            var row = new Row();
            foreach (var columnDefinition in sheetDefinition.ColumnDefinitions)
            {
                var propValue = GetPropertyValue(model, columnDefinition.PropertyName);
                CellValues cellType;
                string cellValue;
                _converterManager.GetCellTypeAndValue(propValue, out cellType, out cellValue);
                var cell = new Cell
                           {
                               DataType = cellType,
                               CellValue = new CellValue(cellValue)
                           };
                cell.AddMdsolAttribute("type",
                    propValue == null ? typeof(object).FullName : propValue.GetType().FullName);
                cell.AddMdsolAttribute("propertyName", columnDefinition.PropertyName);
                row.AppendChild(cell);
            }

            return row;
        }

        private object GetPropertyValue(SheetModel model, string propertyName)
        {
            var value = model.GetPropertyValue(propertyName);
            if (value == null)
            {
                model.GetExtraProperties().TryGetValue(propertyName, out value);
            }
            return value;
        }

        private void BuildSheetFunc(IEnumerable<SheetModel> models, ISheetDefinition sheetDefinition,
                                    SpreadsheetDocument doc)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            var sheets = doc.WorkbookPart.Workbook.Sheets ?? doc.WorkbookPart.Workbook.AppendChild(new Sheets());

            var worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = CreateWorksheet(models, sheetDefinition);

            var sheetId = 1 + sheets.Count();
            var sheetName = sheetDefinition.Name;
            var sheet = new Sheet
                        {
                            Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                            SheetId = (uint) sheetId,
                            Name = sheetName
                        };

            sheets.Append(sheet);
        }

        private Worksheet CreateWorksheet(IEnumerable<SheetModel> models, ISheetDefinition sheetDefinition)
        {
            var sheetData = CreateSheetData(models, sheetDefinition);
            return new Worksheet(sheetData);
        }

        private SheetData CreateSheetData(IEnumerable<SheetModel> models, ISheetDefinition sheetDefinition)
        {
            var sheetData = new SheetData();
            var rows = models.Select(x => BuildRow(x, sheetDefinition));
            sheetData.Append(rows);

            return sheetData;
        }
    }
}
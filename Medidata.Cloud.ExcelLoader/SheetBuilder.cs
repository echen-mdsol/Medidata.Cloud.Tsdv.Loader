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
        private readonly ICellTypeValueConverterFactory _converterFactory;

        public SheetBuilder(ICellTypeValueConverterFactory converterFactory)
        {
            if (converterFactory == null) throw new ArgumentNullException("converterFactory");
            _converterFactory = converterFactory;
            BuildSheet = BuildSheetFunc;
            BuildRow = BuildRowFromExpandoObject;
        }

        public Action<IEnumerable<SheetDefinitionModelBase>, ISheetDefinition, SpreadsheetDocument> BuildSheet { get;
            set; }

        public Func<SheetDefinitionModelBase, ISheetDefinition, Row> BuildRow { get; set; }


        private Row BuildRowFromExpandoObject(SheetDefinitionModelBase model, ISheetDefinition sheetDefinition)
        {
            var row = new Row();
            foreach (var columnDefinition in sheetDefinition.ColumnDefinitions)
            {
                var propValue = GetPropertyValue(model, columnDefinition.PropertyName);
                var converter = _converterFactory.Produce(columnDefinition.PropertyType);
                var cellValue = converter.GetCellValue(propValue);
                var cell = new Cell
                           {
                               DataType = converter.CellType,
                               CellValue = new CellValue(cellValue)
                           };
                row.AppendChild(cell);
            }

            return row;
        }


        private object GetPropertyValue(SheetDefinitionModelBase model, string propertyName)
        {
            var value = model.GetPropertyValue(propertyName);
            if (value == null)
            {
                model.ExtraProperties.TryGetValue(propertyName, out value);
            }
            return value;
        }


        private void BuildSheetFunc(IEnumerable<SheetDefinitionModelBase> models, ISheetDefinition sheetDefinition,
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

            // Use this attribute to retrieve the worksheet.
            var attribute = SpreadsheetAttributeHelper.CreateSheetNameAttribute(sheetName);
            sheet.SetAttribute(attribute);

            sheets.Append(sheet);
        }

        private Worksheet CreateWorksheet(IEnumerable<SheetDefinitionModelBase> models, ISheetDefinition sheetDefinition)
        {
            var sheetData = CreateSheetData(models, sheetDefinition);
            var columns = CreateColumns(sheetData);
            return columns.Any() ? new Worksheet(columns, sheetData) : new Worksheet(sheetData);
        }

        private SheetData CreateSheetData(IEnumerable<SheetDefinitionModelBase> models, ISheetDefinition sheetDefinition)
        {
            var sheetData = new SheetData();
            var rows = models.Select(x => BuildRow(x, sheetDefinition));
            sheetData.Append(rows);

            return sheetData;
        }

        private Columns CreateColumns(SheetData sheetData)
        {
            var numberOfColumns = 0;
            if (sheetData.Descendants<Row>().Any())
            {
                numberOfColumns = sheetData.Descendants<Row>().First().Descendants<Cell>().Count();
            }
            var columnRange = Enumerable.Range(0, numberOfColumns).Select(i => new Column
                                                                               {
                                                                                   Min = (uint) (i + 1),
                                                                                   Max = (uint) (i + 1),
                                                                                   Width = 20D,
                                                                                   CustomWidth = true
                                                                               });
            var cs = new Columns();
            cs.Append(columnRange);
            return cs;
        }
    }
}
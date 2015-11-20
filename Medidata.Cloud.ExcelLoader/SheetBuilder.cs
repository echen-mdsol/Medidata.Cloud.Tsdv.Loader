using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImpromptuInterface;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    internal class SheetBuilder<T> : List<object>, ISheetBuilder where T : class
    {
        private readonly ICellTypeValueConverterFactory _converterFactory;

        public SheetBuilder(ICellTypeValueConverterFactory _converterFactory)
        {
            this._converterFactory = _converterFactory;
        }

        public bool HasHeaderRow { get; set; }
        public string SheetName { get; set; }

        public string[] ColumnNames { get; set; }

        public void AttachTo(SpreadsheetDocument doc)
        {
            var sheets = doc.WorkbookPart.Workbook.Sheets ?? doc.WorkbookPart.Workbook.AppendChild(new Sheets());

            var worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = CreateWorksheet();

            var sheetId = (uint) (1 + sheets.Count());
            var sheet = new Sheet
            {
                Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = SheetName
            };

            // Use this attribute to retrieve the worksheet.
            var attribute = SpreadsheetAttributeHelper.CreateSheetNameAttribute(SheetName);
            sheet.SetAttribute(attribute);

            sheets.Append(sheet);
        }

        private Worksheet CreateWorksheet()
        {
            var sheetData = GetSheetData();
            var worksheet = new Worksheet(sheetData);
            return worksheet;
        }

        private Cell CreateCell(T model, PropertyDescriptor property)
        {
            var cell = new Cell();
            if (model != null)
            {
                var propValue = GetPropertyValue(property, model);
                var converter = _converterFactory.Produce(property.PropertyType);
                var cellType = converter.CellType;
                var cellValue = converter.GetCellValue(propValue);
                cell.DataType = cellType;
                if (cellType == CellValues.InlineString)
                {
                    cell.InlineString = new InlineString {Text = new Text(cellValue)};
                }
                else
                {
                    cell.CellValue = new CellValue(cellValue);
                }
            }
            return cell;
        }

        private object GetPropertyValue(PropertyDescriptor property, object target)
        {
            try
            {
                return property.GetValue(target);
            }
            catch
            {
                return null;
            }
        }

        private Row CreateRow(T model)
        {
            var row = new Row();
            var properties = model.GetType().GetPropertyDescriptors();
            foreach (var prop in properties)
            {
                var cell = CreateCell(model, prop);
                row.AppendChild(cell);
            }
            return row;
        }

        private Row CreateHeaderRow()
        {
            var row = new Row();
            foreach (var columnName in ColumnNames)
            {
                var cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(columnName)
                };
                row.AppendChild(cell);
            }
            return row;
        }

        private SheetData GetSheetData()
        {
            var sheetData = new SheetData();
            if (HasHeaderRow)
            {
                var headerRow = CreateHeaderRow();
                sheetData.Append(headerRow);
            }

            var rows = this.Select(x => (x as T) ?? x.ActLike<T>()).Select(CreateRow);
            sheetData.Append(rows);

            return sheetData;
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImpromptuInterface;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    internal class SheetBuilder<T> : List<T>, ISheetBuilder where T : class
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
            var columns = GetColumns(sheetData);
            Worksheet worksheet;
            if (columns != null && columns.Any())
            {
                worksheet = new Worksheet(columns, sheetData);
            }
            else
            {
                worksheet = new Worksheet(sheetData);
            }
            
            return worksheet;
        }

        private Columns GetColumns(SheetData sheetData)
        {
            int numberOfColumns = 0;
            if (sheetData.Descendants<Row>().Any())
            {
                numberOfColumns = sheetData.Descendants<Row>().First().Descendants<Cell>().Count();
            }
            Columns cs = new Columns();
            for (int i = 0; i < numberOfColumns; i++)
            {
                Column c = new Column()
                {
                    Min = (uint)(i+1),
                    Max = (uint)(i+1),
                    Width = 20D,
                    CustomWidth = true
                };
                cs.Append(c);
            }
            return cs;

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
                    //var t = property.PropertyType;
                    //if (t == typeof (DateTime))
                    //{
                    //    cell.StyleIndex = 164;    
                    //}
                    //if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof (Nullable<>) &&
                    //    Nullable.GetUnderlyingType(t) == typeof (DateTime))
                    //{
                    //    cell.StyleIndex = 164;    
                    //}
                    
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

            var rows = this.Select(x => x ?? x.ActLike<T>()).Select(CreateRow);
            sheetData.Append(rows);
            
            return sheetData;
        }

    }
}
using System.Collections.Generic;
using System.ComponentModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImpromptuInterface;
using Medidata.Cloud.Tsdv.Loader.CellValueConverters;
using Medidata.Cloud.Tsdv.Loader.Extensions;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class WorksheetBuilder<T> : List<object>, IWorksheetBuilder where T : class
    {
        private readonly CellTypeConverterManager _converterManager;
        private bool _hasHeaderRow;

        public WorksheetBuilder()
        {
            _converterManager = new CellTypeConverterManager();
        }

        public string[] ColumnNames { get; set; }

        public void AppendWorksheet(SpreadsheetDocument doc, bool hasHeaderRow, string sheetName)
        {
            _hasHeaderRow = hasHeaderRow;
            var worksheet = CreateWorksheet();
            WorksheetBuilderHelper.AppendWorksheet(doc, worksheet, sheetName);
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
                CellValues cellType;
                string cellValue;
                var propValue = property.GetValue(model);
                _converterManager.GetCellTypeAndValue(property.PropertyType, propValue, out cellType, out cellValue);
                cell.DataType = cellType;
                cell.CellValue = new CellValue(cellValue);
//                cell.SetAttribute(new OpenXmlAttribute("PropertyName", "http://www.msdol.com", property.Name));
            }
            return cell;
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
            if (_hasHeaderRow)
            {
                var headerRow = CreateHeaderRow();
                sheetData.Append(headerRow);
            }

            foreach (var element in this)
            {
                T model = element.ActLike<T>();
                var row = CreateRow(model);
                sheetData.Append(row);
            }

            return sheetData;
        }
    }
}
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.Tsdv.Loader.CellValueConverters;
using Medidata.Cloud.Tsdv.Loader.Extensions;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class WorksheetBuilder<T> : List<T>, IWorksheetBuilder
    {
        private readonly Type _objectType;
        private readonly CellTypeConverterManager _converterManager;

        public WorksheetBuilder()
        {
            _objectType = typeof(T);
            _converterManager = new CellTypeConverterManager();
        }

        public string[] ColumnNames { get; set; }

        public void AppendWorksheet(SpreadsheetDocument doc, string sheetName)
        {
            var worksheet = CreateWorksheet();
            WorksheetBuilderHelper.AppendWorksheet(doc, worksheet, sheetName);
        }

        private Worksheet CreateWorksheet()
        {
            var sheetData = GetSheetData();
            var worksheet = new Worksheet(sheetData);
            return worksheet;
        }

        private Cell CreateCell(object model, PropertyDescriptor property)
        {
            var cell = new Cell();
            if (model != null)
            {
                CellValues cellType;
                string cellValue;
                var propValue = property.GetValue(model);
                _converterManager.GetCellType(property.PropertyType, propValue, out cellType, out cellValue);
                cell.DataType = cellType;
                cell.CellValue = new CellValue(cellValue);
                cell.SetAttribute(new OpenXmlAttribute("PropertyName", "http://www.msdol.com", property.Name));
            }
            return cell;
        }

        private Row CreateRow(object model)
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
            var headerRow = CreateHeaderRow();
            sheetData.Append(headerRow);
            foreach (var model in this)
            {
                var row = CreateRow(model);
                sheetData.Append(row);
            }

            return sheetData;
        }
    }
}
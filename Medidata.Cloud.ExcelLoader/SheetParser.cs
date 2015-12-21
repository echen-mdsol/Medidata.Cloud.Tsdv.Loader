using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public class SheetParser : ISheetParser
    {
        private readonly ICellTypeValueConverterFactory _converterFactory;

        public SheetParser(ICellTypeValueConverterFactory converterFactory)
        {
            if (converterFactory == null) throw new ArgumentNullException("converterFactory");
            _converterFactory = converterFactory;
        }

        public IEnumerable<ExpandoObject> GetObjects(Worksheet worksheet, ISheetDefinition sheetDefinition)
        {
            if (worksheet == null) throw new ArgumentNullException("worksheet");
            if (sheetDefinition == null) throw new ArgumentNullException("sheetDefinition");
            var headerRow = worksheet.Descendants<Row>().First();
            var rows = worksheet.Descendants<Row>().Skip(1);
            return rows.Select(x => ParseFromRow(x, sheetDefinition, headerRow));
        }

        private ExpandoObject ParseFromRow(Row row, ISheetDefinition sheetDefinition, Row headerRow)
        {
            IDictionary<string, object> expando = new ExpandoObject();
            var cells = row.Elements<Cell>().ToList();
            var colDefs = sheetDefinition.ColumnDefinitions.ToList();
            var count = Math.Max(cells.Count, colDefs.Count);
            for (var index = 0; index < count; index++)
            {
                var colDef = index < colDefs.Count ? colDefs[index] : null;
                var cell = index < cells.Count ? cells[index] : null;
                string propName;
                object propValue;
                if (cell != null && colDef != null)
                {
                    ParsePropNameAndValue(cell, colDef, out propName, out propValue);
                }
                else if (colDef != null)
                {
                    ParsePropNameAndValue(colDef, out propName, out propValue);
                }
                else if (cell != null)
                {
                    var headerCell = headerRow.Elements<Cell>().ElementAt(index);
                    ParsePropNameAndValue(cell, headerCell, out propName, out propValue);
                }
                else
                {
                    throw new NotSupportedException("Cell and CellDefinition are both null");
                }

                expando.Add(propName, propValue);
            }
            return (ExpandoObject)expando;
        }

        private void ParsePropNameAndValue(Cell cell, IColumnDefinition colDef, out string propName, out object propValue)
        {
            var converter = _converterFactory.Produce(colDef.PropertyType);
            propName = colDef.PropertyName;
            propValue = converter.GetCSharpValue(cell.InnerText);
        }

        private void ParsePropNameAndValue(IColumnDefinition colDef, out string propName, out object propValue)
        {
            propName = colDef.PropertyName;
            propValue = null;
        }

        private void ParsePropNameAndValue(Cell cell, Cell headerCell, out string propName, out object propValue)
        {
            var converter = _converterFactory.Produce(cell.DataType);
            propName = headerCell.InnerText;
            propValue = converter.GetCSharpValue(cell.InnerText);
        }
    }
}
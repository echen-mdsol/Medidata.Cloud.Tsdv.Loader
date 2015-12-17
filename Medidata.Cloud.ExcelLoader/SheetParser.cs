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
            var rows = worksheet.Descendants<Row>().Skip(1);
            return rows.Select(x => ParseFromRow(x, sheetDefinition));
        }

        private ExpandoObject ParseFromRow(Row row, ISheetDefinition sheetDefinition)
        {
            IDictionary<string, object> expando = new ExpandoObject();
            var cells = row.Elements<Cell>().ToList();
            var index = 0;
            foreach (var colDef in sheetDefinition.ColumnDefinitions)
            {
                var converter = _converterFactory.Produce(colDef.PropertyType);
                var propValue = converter.GetCSharpValue(cells[index].InnerText);
                expando.Add(colDef.PropertyName, propValue);
                index ++;
            }
            return (ExpandoObject) expando;
        }
    }
}
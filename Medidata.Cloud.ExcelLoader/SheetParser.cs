using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    internal class SheetParser : ISheetParser
    {
        private readonly ICellTypeValueConverterFactory _converterFactory;
        private Worksheet _worksheet;

        public SheetParser(ICellTypeValueConverterFactory converterFactory)
        {
            if (converterFactory == null) throw new ArgumentNullException("converterFactory");
            _converterFactory = converterFactory;
        }

        public IEnumerable<ExpandoObject> GetObjects(ISheetDefinition sheetDefinition)
        {
            var rows = _worksheet.Descendants<Row>().Skip(1);
            return rows.Select(x => ParseFromRow(x, sheetDefinition));
        }

        public void Load(Worksheet worksheet)
        {
            _worksheet = worksheet;
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
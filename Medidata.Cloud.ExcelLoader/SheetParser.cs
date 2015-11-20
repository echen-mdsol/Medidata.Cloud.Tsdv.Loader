using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ImpromptuInterface;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    internal class SheetParser<T> : ISheetParser<T> where T : class
    {
        private readonly ICellTypeValueConverterFactory _converterFactory;
        private Worksheet _worksheet;

        public SheetParser(ICellTypeValueConverterFactory converterFactory)
        {
            if (converterFactory == null) throw new ArgumentNullException("converterFactory");
            _converterFactory = converterFactory;
        }

        public bool HasHeaderRow { get; set; }

        public IEnumerable<T> GetObjects()
        {
            var rows = _worksheet.Descendants<Row>().Skip(HasHeaderRow ? 1 : 0);
            return rows.Select(ParseFromRow);
        }

        public void Load(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        private T ParseFromRow(Row row)
        {
            var properties = typeof (T).GetPropertyDescriptors();
            IDictionary<string, object> expando = new ExpandoObject();
            var cells = row.Elements<Cell>().ToList();
            var index = 0;
            foreach (var prop in properties)
            {
                var converter = _converterFactory.Produce(prop.PropertyType);
                var propValue = converter.GetCSharpValue(cells[index].InnerText);
                expando.Add(prop.Name, propValue);
                index ++;
            }
            T target = expando.ActLike();
            return target;
        }
    }
}
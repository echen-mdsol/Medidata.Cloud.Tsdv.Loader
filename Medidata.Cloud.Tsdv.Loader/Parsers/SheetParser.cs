using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ImpromptuInterface;
using Medidata.Cloud.Tsdv.Loader.CellTypeConverters;
using Medidata.Cloud.Tsdv.Loader.Extensions;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    internal class SheetParser<T> : ISheetParser<T> where T : class
    {
        private readonly CellTypeValueConverterManager _converter;
        private Worksheet _worksheet;

        public SheetParser()
        {
            _converter = new CellTypeValueConverterManager();
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
                var cell = cells[index];
                var propValue = _converter.GetCSharpValue(prop.PropertyType, cell.DataType, cell.CellValue.InnerText);
                expando.Add(prop.Name, propValue);
                index ++;
            }
            T target = expando.ActLike();
            return target;
        }
    }
}
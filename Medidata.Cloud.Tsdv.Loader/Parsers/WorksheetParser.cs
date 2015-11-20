using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ImpromptuInterface;
using Medidata.Cloud.Tsdv.Loader.CellValueConverters;
using Medidata.Cloud.Tsdv.Loader.Extensions;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public class WorksheetParser<T> : IWorksheetParser<T> where T : class
    {
        private readonly CellTypeConverterManager _converterManager;
        private bool _hasHeaderRow;
        private Worksheet _worksheet;

        public WorksheetParser()
        {
            _converterManager = new CellTypeConverterManager();
        }

        public IEnumerable<T> GetObjects()
        {
            var rows = _worksheet.Descendants<Row>().Skip(_hasHeaderRow ? 1 : 0);
            return rows.Select(ParseFromRow);
        }

        public void Load(Worksheet worksheet, bool hasHeaderRow = true)
        {
            _worksheet = worksheet;
            _hasHeaderRow = hasHeaderRow;
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
                object propValue;
                _converterManager.GetCSharpValue(cell.DataType, cell.CellValue.InnerText, prop.PropertyType,
                    out propValue);
                expando.Add(prop.Name, propValue);
                index ++;
            }
            T target = expando.ActLike();
            return target;
        }
    }
}
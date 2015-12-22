using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    public class CellTypeValueConverterManager : ICellTypeValueConverterManager
    {
        private readonly IEnumerable<ICellTypeValueConverter> _converters;

        public CellTypeValueConverterManager() : this(null) {}

        public CellTypeValueConverterManager(params ICellTypeValueConverter[] converters)
        {
            _converters = new ICellTypeValueConverter[]
                          {
                              new BooleanConverter(),
                              new NullableBooleanConverter(),
                              new IntConverter(),
                              new DoubleConverter(),
                              new FloatConverter(),
                              new DecimalConverter(),
                              new DateTimeConverter(),
                              new NullableDateTimeConverter(),
                              new StringConverter()
                          };
            if (converters != null)
            {
                _converters = _converters.Concat(converters);
            }
        }

        public void GetCellTypeAndValue(object value, out CellValues cellType, out string cellValue)
        {
            string parsedCellValue = null;
            var converter = _converters.FirstOrDefault(x => x.TryConvertToCellValue(value, out parsedCellValue));
            if (converter == null)
            {
                var msg =
                    string.Format("Cannot find a propery converter to convert the CSharp value to cell. Value: '{0}'",
                        value);
                throw new NotSupportedException(msg);
            }

            cellValue = parsedCellValue;
            cellType = converter.CellType;
        }

        public object GetCSharpValue(Cell cell)
        {
            object value = null;
            var cellValue = cell.InnerText;
            var converter = _converters.FirstOrDefault(x => x.TryConvertToCSharpValue(cellValue, out value));
            if (converter == null)
            {
                var msg =
                    string.Format("Cannot find a propery converter to parse the cell to CSharp value. Value: '{0}'",
                        cellValue);
                throw new NotSupportedException(msg);
            }

            var propType = Type.GetType(cell.GetMdsolAttribute("type"), true);
            var propValue = Convert.ChangeType(value, propType);
            return propValue;
        }
    }
}
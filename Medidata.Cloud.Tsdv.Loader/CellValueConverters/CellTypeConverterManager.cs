using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class CellTypeConverterManager
    {
        private readonly ICellValueConverter[] _converters;

        public CellTypeConverterManager(ICellValueConverter[] converters = null)
        {
            _converters = new ICellValueConverter[]
            {
                new BooleanConverter(),
                new NullableBooleanConverter(),
                new DateTimeConverter(),
                new NullableDateTimeConverter(),
                new DoubleConverter(),
                new FloatConverter(),
                new DecimalConverter(),
                new IntConverter(),
                new LongConverter(),
                new StringConverter(),
                new InlineStringConverter()
            };

            if (converters != null)
            {
                _converters = _converters.Concat(converters).ToArray();
            }
        }

        public void GetCellType(Type type, object value, out CellValues cellType, out string cellValue)
        {
            var converter = _converters.First(c => c.CSharpType == type);
            cellType = converter.CellType;
            cellValue = converter.GetCellValue(value);
        }

        public void GetCSharpType(CellValues cellType, string cellValue, Type type, out object value)
        {
            var converter = _converters.First(c => c.CSharpType == type && c.CellType == cellType);
            value = converter.GetCSharpValue(cellValue);
        }
    }
}
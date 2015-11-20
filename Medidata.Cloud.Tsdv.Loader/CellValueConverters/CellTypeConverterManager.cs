using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class CellTypeConverterManager
    {
        private readonly ICellValueConverter[] _converters;

        public CellTypeConverterManager() : this(null)
        {
        }

        public CellTypeConverterManager(ICellValueConverter[] customConverters)
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

            if (customConverters != null)
            {
                _converters = _converters.Concat(customConverters).ToArray();
            }
        }

        public void GetCellTypeAndValue(Type type, object value, out CellValues cellType, out string cellValue)
        {
            var converter = _converters.First(c => c.CSharpType == type);
            cellType = converter.CellType;
            cellValue = converter.GetCellValue(value);
        }

        public void GetCSharpValue(CellValues cellType, string cellValue, Type type, out object value)
        {
            var converter = _converters.First(c => c.CSharpType == type && c.CellType == cellType);
            value = converter.GetCSharpValue(cellValue);
        }
    }
}
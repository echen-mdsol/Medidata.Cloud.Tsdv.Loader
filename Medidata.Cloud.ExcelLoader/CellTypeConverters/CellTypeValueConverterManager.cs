using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class CellTypeValueConverterManager
    {
        private readonly ICellTypeValueConverter[] _converters;

        public CellTypeValueConverterManager() : this(null)
        {
        }

        public CellTypeValueConverterManager(ICellTypeValueConverter[] customConverters)
        {
            _converters = new ICellTypeValueConverter[]
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

        public object GetCSharpValue(Type type, CellValues cellType, string cellValue)
        {
            var converter = _converters.First(c => c.CSharpType == type && c.CellType == cellType);
            return converter.GetCSharpValue(cellValue);
        }
    }
}
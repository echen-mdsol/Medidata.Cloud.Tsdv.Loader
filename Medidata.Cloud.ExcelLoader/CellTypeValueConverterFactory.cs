using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;

namespace Medidata.Cloud.ExcelLoader
{
    public class CellTypeValueConverterFactory : ICellTypeValueConverterFactory
    {
        private readonly IEnumerable<ICellTypeValueConverter> _converters;

        public CellTypeValueConverterFactory() : this(null) {}

        public CellTypeValueConverterFactory(params ICellTypeValueConverter[] converters)
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
                                                                         new StringConverter()
                                                                     };
            if (converters != null)
            {
                _converters = _converters.Concat(converters);
            }
        }

        public ICellTypeValueConverter Produce(Type type)
        {
            return _converters.First(x => x.CSharpType == type);
        }

        public ICellTypeValueConverter Produce(CellValues cellType)
        {
            return _converters.First(x => x.CellType == cellType);
        }
    }
}
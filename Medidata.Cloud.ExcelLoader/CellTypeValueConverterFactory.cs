using System;
using System.Collections.Generic;
using System.Linq;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;

namespace Medidata.Cloud.ExcelLoader
{
    public class CellTypeValueConverterFactory : ICellTypeValueConverterFactory
    {
        private readonly IDictionary<Type, ICellTypeValueConverter> _converters;

        public CellTypeValueConverterFactory() : this(null)
        {
        }

        public CellTypeValueConverterFactory(ICellTypeValueConverter[] customConverters)
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
            }.ToDictionary(x => x.CSharpType, x => x);

            if (customConverters != null)
            {
                foreach (var customConverter in customConverters)
                {
                    var type = customConverter.CSharpType;
                    if (_converters.ContainsKey(type))
                    {
                        _converters[type] = customConverter;
                    }
                    else
                    {
                        _converters.Add(type, customConverter);
                    }
                }
            }
        }

        public ICellTypeValueConverter Produce(Type type)
        {
            return _converters[type];
        }
    }
}
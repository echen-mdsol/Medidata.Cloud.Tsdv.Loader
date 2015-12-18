using System;
using System.Collections.Generic;
using System.Linq;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;

namespace Medidata.Cloud.ExcelLoader
{
    public class CellTypeValueConverterFactory : ICellTypeValueConverterFactory
    {
        private readonly IDictionary<Type, ICellTypeValueConverter> _converters;

        public CellTypeValueConverterFactory() : this(null) {}

        public CellTypeValueConverterFactory(params ICellTypeValueConverter[] converters)
        {
            IEnumerable<ICellTypeValueConverter> defaultConverters = new ICellTypeValueConverter[]
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
                defaultConverters = defaultConverters.Concat(converters);
            }
            _converters = defaultConverters.ToDictionary(x => x.CSharpType, x => x);
        }

        public ICellTypeValueConverter Produce(Type type)
        {
            return _converters[type];
        }
    }
}
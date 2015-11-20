using System;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ICellTypeValueConverterFactory
    {
        ICellTypeValueConverter Produce(Type type);
    }
}
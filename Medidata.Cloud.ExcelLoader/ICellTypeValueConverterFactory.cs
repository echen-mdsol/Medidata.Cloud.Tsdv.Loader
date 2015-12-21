using System;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.CellTypeConverters;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ICellTypeValueConverterFactory
    {
        ICellTypeValueConverter Produce(Type type);
        ICellTypeValueConverter Produce(CellValues cellType);
    }
}
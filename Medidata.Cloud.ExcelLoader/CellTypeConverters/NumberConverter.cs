using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal abstract class NumberConverter<T> : CellTypeValueBaseConverter<T>
    {
        protected NumberConverter()
            : base(CellValues.Number)
        {
        }
    }
}
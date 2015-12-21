using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal abstract class NumberConverter<T> : CellTypeValueBaseConverter<T>
    {
        protected NumberConverter()
            : base(CellValues.Number) {}

        protected override bool TryGetCellValueImpl(T csharpValue, out string cellValue)
        {
            cellValue = csharpValue.ToString();
            return true;
        }
    }
}
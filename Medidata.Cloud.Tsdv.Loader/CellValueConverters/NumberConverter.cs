using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal abstract class NumberConverter<T> : CellValueBaseConverter<T>
    {
        protected NumberConverter()
            : base(CellValues.Number)
        {
        }

        protected override string GetCellValueImpl(T csharpValue)
        {
            return csharpValue == null ? null : csharpValue.ToString();
        }
    }
}
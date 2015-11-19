using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class NullableBooleanConverter : CellValueBaseConverter<bool?>
    {
        public NullableBooleanConverter()
            : base(CellValues.Boolean)
        {
        }

        protected override string GetCellValueImpl(bool? csharpValue)
        {
            return csharpValue.HasValue && csharpValue.Value ? "1" : "0";
        }

        protected override bool? GetCSharpValueImpl(string cellValue)
        {
            return string.IsNullOrWhiteSpace(cellValue) ? (bool?) null : bool.Parse(cellValue);
        }
    }
}
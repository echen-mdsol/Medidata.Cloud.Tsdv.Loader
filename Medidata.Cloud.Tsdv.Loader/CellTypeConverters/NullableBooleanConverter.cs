using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellTypeConverters
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
            if (cellValue == "1") return true;
            if (cellValue == "0") return false;
            bool value;
            return bool.TryParse(cellValue, out value) ? value : (bool?) null;
        }
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class NullableBooleanConverter : CellTypeValueBaseConverter<bool?>
    {
        public NullableBooleanConverter()
            : base(CellValues.Boolean) {}

        protected override string GetCellValueImpl(bool? csharpValue)
        {
            return csharpValue.HasValue ? BooleanValue.FromBoolean(csharpValue.Value) : BooleanValue.FromBoolean(false);
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
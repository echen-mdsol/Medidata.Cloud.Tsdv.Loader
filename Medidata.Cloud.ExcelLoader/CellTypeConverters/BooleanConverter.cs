using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class BooleanConverter : CellTypeValueBaseConverter<bool>
    {
        public BooleanConverter()
            : base(CellValues.Boolean)
        {
        }

        protected override string GetCellValueImpl(bool csharpValue)
        {
            return csharpValue ? "1" : "0";
        }

        protected override bool GetCSharpValueImpl(string cellValue)
        {
            if (cellValue == "1") return true;
            if (cellValue == "0") return false;
            bool value;
            bool.TryParse(cellValue, out value);
            return value;
        }
    }
}
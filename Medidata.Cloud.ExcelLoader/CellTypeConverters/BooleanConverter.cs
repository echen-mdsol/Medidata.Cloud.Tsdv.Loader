using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class BooleanConverter : CellTypeValueBaseConverter<bool>
    {
        public BooleanConverter()
            : base(CellValues.Boolean) {}

        protected override bool TryGetCellValueImpl(bool csharpValue, out string cellValue)
        {
            cellValue = csharpValue ? "1" : "0";
            return true;
        }

        protected override bool TryGetCSharpValueImpl(string cellValue, out bool csharpValue)
        {
            if (cellValue == "1")
            {
                csharpValue = true;
                return true;
            }
            if (cellValue == "0")
            {
                csharpValue = false;
                return true;
            }
            return bool.TryParse(cellValue, out csharpValue);
        }
    }
}
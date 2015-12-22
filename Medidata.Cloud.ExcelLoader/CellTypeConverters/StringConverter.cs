using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class StringConverter : CellTypeValueBaseConverter<string>
    {
        public StringConverter() : base(CellValues.String) {}

        protected override bool TryGetCellValueImpl(string csharpValue, out string cellValue)
        {
            cellValue = csharpValue;
            return true;
        }

        protected override bool TryGetCSharpValueImpl(string cellValue, out string csharpValue)
        {
            csharpValue = cellValue;
            return true;
        }
    }
}
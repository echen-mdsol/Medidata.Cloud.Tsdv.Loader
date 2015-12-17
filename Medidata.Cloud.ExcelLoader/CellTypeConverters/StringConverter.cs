using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class StringConverter : CellTypeValueBaseConverter<string>
    {
        public StringConverter() : base(CellValues.String) {}

        protected override string GetCellValueImpl(string csharpValue)
        {
            return csharpValue;
        }

        protected override string GetCSharpValueImpl(string cellValue)
        {
            return cellValue;
        }
    }
}
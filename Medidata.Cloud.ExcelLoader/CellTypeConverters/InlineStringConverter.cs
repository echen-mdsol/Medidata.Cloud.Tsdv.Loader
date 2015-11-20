using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class InlineStringConverter : CellTypeValueBaseConverter<string>
    {
        public InlineStringConverter()
            : base(CellValues.InlineString)
        {
        }

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
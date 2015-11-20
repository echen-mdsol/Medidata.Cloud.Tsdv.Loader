using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellTypeConverters
{
    internal class InlineStringConverter : CellValueBaseConverter<string>
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
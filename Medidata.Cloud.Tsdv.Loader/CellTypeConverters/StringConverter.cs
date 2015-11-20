using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellTypeConverters
{
    internal class StringConverter : CellValueBaseConverter<string>
    {
        public StringConverter() : base(CellValues.String)
        {
        }

        protected override string GetCellValueImpl(string csharpValue)
        {
            return csharpValue;
            ;
        }

        protected override string GetCSharpValueImpl(string cellValue)
        {
            return cellValue;
        }
    }
}
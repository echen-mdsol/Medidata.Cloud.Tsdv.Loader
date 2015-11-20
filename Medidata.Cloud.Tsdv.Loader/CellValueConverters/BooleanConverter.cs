using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class BooleanConverter : CellValueBaseConverter<bool>
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
            return cellValue == "0" || cellValue != "1" && bool.Parse(cellValue);
        }
    }
}
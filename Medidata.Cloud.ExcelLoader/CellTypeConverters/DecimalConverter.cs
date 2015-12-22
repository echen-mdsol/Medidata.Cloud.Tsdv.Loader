namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class DecimalConverter : NumberConverter<decimal>
    {
        protected override bool TryGetCSharpValueImpl(string cellValue, out decimal csharpValue)
        {
            return decimal.TryParse(cellValue, out csharpValue);
        }
    }
}
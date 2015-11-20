namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class DecimalConverter : NumberConverter<decimal>
    {
        protected override decimal GetCSharpValueImpl(string cellValue)
        {
            decimal value;
            decimal.TryParse(cellValue, out value);
            return value;
        }
    }
}
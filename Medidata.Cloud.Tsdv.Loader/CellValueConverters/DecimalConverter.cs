namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class DecimalConverter : NumberConverter<decimal>
    {
        protected override decimal GetCSharpValueImpl(string cellValue)
        {
            return int.Parse(cellValue);
        }
    }
}
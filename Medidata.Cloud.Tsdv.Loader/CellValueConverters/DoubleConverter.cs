namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class DoubleConverter : NumberConverter<double>
    {
        protected override double GetCSharpValueImpl(string cellValue)
        {
            return double.Parse(cellValue);
        }
    }
}
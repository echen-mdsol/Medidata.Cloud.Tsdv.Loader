namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class DoubleConverter : NumberConverter<double>
    {
        protected override bool TryGetCSharpValueImpl(string cellValue, out double csharpValue)
        {
            return double.TryParse(cellValue, out csharpValue);
        }
    }
}
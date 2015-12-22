namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class FloatConverter : NumberConverter<float>
    {
        protected override bool TryGetCSharpValueImpl(string cellValue, out float csharpValue)
        {
            return float.TryParse(cellValue, out csharpValue);
        }
    }
}
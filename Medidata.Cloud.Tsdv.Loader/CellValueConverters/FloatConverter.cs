namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class FloatConverter : NumberConverter<float>
    {
        protected override float GetCSharpValueImpl(string cellValue)
        {
            return float.Parse(cellValue);
        }
    }
}
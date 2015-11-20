namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class FloatConverter : NumberConverter<float>
    {
        protected override float GetCSharpValueImpl(string cellValue)
        {
            float value;
            float.TryParse(cellValue, out value);
            return value;
        }
    }
}
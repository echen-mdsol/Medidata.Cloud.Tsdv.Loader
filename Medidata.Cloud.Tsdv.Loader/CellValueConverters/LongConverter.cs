namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class LongConverter : NumberConverter<long>
    {
        protected override long GetCSharpValueImpl(string cellValue)
        {
            long value;
            long.TryParse(cellValue, out value);
            return value;
        }
    }
}
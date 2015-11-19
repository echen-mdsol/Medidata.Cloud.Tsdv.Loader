namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class LongConverter : NumberConverter<long>
    {
        protected override long GetCSharpValueImpl(string cellValue)
        {
            return long.Parse(cellValue);
        }
    }
}
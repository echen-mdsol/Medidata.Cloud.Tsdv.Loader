namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal class IntConverter : NumberConverter<int>
    {
        protected override int GetCSharpValueImpl(string cellValue)
        {
            int value;
            int.TryParse(cellValue, out value);
            return value;
        }
    }
}
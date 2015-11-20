namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
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
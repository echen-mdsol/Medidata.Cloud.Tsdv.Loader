namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class IntConverter : NumberConverter<int>
    {
        protected override bool TryGetCSharpValueImpl(string cellValue, out int csharpValue)
        {
            if (string.IsNullOrWhiteSpace(cellValue) || cellValue.Contains("."))
            {
                csharpValue = 0;
                return false;
            }

            return int.TryParse(cellValue, out csharpValue);
        }
    }
}
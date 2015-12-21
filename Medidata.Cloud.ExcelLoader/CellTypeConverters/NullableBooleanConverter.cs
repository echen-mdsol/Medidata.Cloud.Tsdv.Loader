using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class NullableBooleanConverter : CellTypeValueBaseConverter<bool?>
    {
        private readonly BooleanConverter _booleanConverter;

        public NullableBooleanConverter()
            : base(CellValues.Boolean)
        {
            _booleanConverter = new BooleanConverter();
        }

        protected override bool TryGetCellValueImpl(bool? csharpValue, out string cellValue)
        {
            cellValue = null;
            if (csharpValue.HasValue)
            {
                return _booleanConverter.TryConvertToCellValue(csharpValue, out cellValue);
            }

            return true;
        }

        protected override bool TryGetCSharpValueImpl(string cellValue, out bool? csharpValue)
        {
            csharpValue = null;
            if (string.IsNullOrEmpty(cellValue))
            {
                return true;
            }

            object value;
            var result = _booleanConverter.TryConvertToCSharpValue(cellValue, out value);
            if (result)
            {
                csharpValue = (bool) value;
                return true;
            }

            return false;
        }
    }
}
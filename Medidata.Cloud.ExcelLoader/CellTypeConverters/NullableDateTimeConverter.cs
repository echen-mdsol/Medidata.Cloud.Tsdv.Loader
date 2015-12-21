using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class NullableDateTimeConverter : CellTypeValueBaseConverter<DateTime?>
    {
        private readonly DateTimeConverter _dateTimeConverter;

        public NullableDateTimeConverter()
            : base(CellValues.String)
        {
            _dateTimeConverter = new DateTimeConverter();
        }

        protected override bool TryGetCellValueImpl(DateTime? csharpValue, out string cellValue)
        {
            cellValue = null;
            if (csharpValue.HasValue)
            {
                return _dateTimeConverter.TryConvertToCellValue(csharpValue, out cellValue);
            }
            return true;
        }

        protected override bool TryGetCSharpValueImpl(string cellValue, out DateTime? csharpValue)
        {
            csharpValue = null;
            if (string.IsNullOrEmpty(cellValue))
            {
                return true;
            }

            object value;
            var result = _dateTimeConverter.TryConvertToCSharpValue(cellValue, out value);
            if (result)
            {
                csharpValue = (DateTime) value;
                return true;
            }

            return false;
        }
    }
}
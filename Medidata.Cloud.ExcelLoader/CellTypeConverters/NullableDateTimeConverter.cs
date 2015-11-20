using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class NullableDateTimeConverter : CellTypeValueBaseConverter<DateTime?>
    {
        public NullableDateTimeConverter()
            : base(CellValues.Date)
        {
        }

        protected override string GetCellValueImpl(DateTime? csharpValue)
        {
            return csharpValue.HasValue ? csharpValue.ToString() : string.Empty;
        }

        protected override DateTime? GetCSharpValueImpl(string cellValue)
        {
            DateTime value;
            return DateTime.TryParse(cellValue, out value) ? value : (DateTime?) null;
        }
    }
}
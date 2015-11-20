using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellTypeConverters
{
    internal class DateTimeConverter : CellValueBaseConverter<DateTime>
    {
        public DateTimeConverter()
            : base(CellValues.Date)
        {
        }

        protected override string GetCellValueImpl(DateTime csharpValue)
        {
            return csharpValue.ToString();
        }

        protected override DateTime GetCSharpValueImpl(string cellValue)
        {
            DateTime value;
            DateTime.TryParse(cellValue, out value);
            return value;
        }
    }
}
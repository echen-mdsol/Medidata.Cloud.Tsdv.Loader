using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class NullableDateTimeConverter : CellTypeValueBaseConverter<DateTime?>
    {
        public NullableDateTimeConverter()
            : base(CellValues.String) {}

        protected override string GetCellValueImpl(DateTime? csharpValue)
        {
            return csharpValue.HasValue ? csharpValue.Value.ToString("MM/dd/yyyy") : string.Empty;
        }

        protected override DateTime? GetCSharpValueImpl(string cellValue)
        {
            DateTime outDt;
            if (DateTime.TryParse(cellValue, out outDt))
            {
                return outDt;
            }
            //Check and handle OLE Automation date time
            if (IsNumeric(cellValue))
            {
                double outNum = 0;
                double.TryParse(cellValue, out outNum);
                return DateTime.FromOADate(outNum);
            }
            return null;
        }
    }
}
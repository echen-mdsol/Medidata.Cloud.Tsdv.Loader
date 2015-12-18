using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class DateTimeConverter : CellTypeValueBaseConverter<DateTime>
    {
        public DateTimeConverter()
            : base(CellValues.Number) {}

        protected override string GetCellValueImpl(DateTime csharpValue)
        {
            //TODO: Discuss if we should use : return csharpValue.ToOADate().ToString(CultureInfo.InvariantCulture);
            return csharpValue.ToString("MM/dd/yyyy");
        }

        protected override DateTime GetCSharpValueImpl(string cellValue)
        {
            DateTime outDt;
            if (DateTime.TryParse(cellValue, out outDt))
            {
                return outDt;
            }
            //Check OLE Automation date time
            if (IsNumeric(cellValue))
            {
                double outNum = 0;
                double.TryParse(cellValue, out outNum);
                return DateTime.FromOADate(outNum);
            }
            return DateTime.Now;
        }
    }
}
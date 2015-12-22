using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal class DateTimeConverter : CellTypeValueBaseConverter<DateTime>
    {
        public DateTimeConverter()
            : base(CellValues.String) {}

        protected override bool TryGetCellValueImpl(DateTime csharpValue, out string cellValue)
        {
            cellValue = csharpValue.ToString("yyyy/MM/dd");
            return true;
        }

        protected override bool TryGetCSharpValueImpl(string cellValue, out DateTime csharpValue)
        {
            DateTime outDt;
            if (DateTime.TryParse(cellValue, out outDt))
            {
                csharpValue = outDt;
                return true;
            }
            //Check OLE Automation date time
            if (IsNumeric(cellValue))
            {
                double outNum;
                double.TryParse(cellValue, out outNum);
                try
                {
                    csharpValue = DateTime.FromOADate(outNum);
                    return true;
                }
                catch
                {
                    // Intentially swallows the exception.
                }
            }

            csharpValue = default(DateTime);
            return false;
        }
    }
}
using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public class HeaderAutoFilterProvider : IAutoFilterProvider
    {
        public AutoFilter CreateAutoFilter(SheetData sheetData)
        {
            var firstRow = sheetData.Descendants<Row>().FirstOrDefault();
            var numberOfColumns = firstRow == null ? 0 : firstRow.Descendants<Cell>().Count();

            //No data, no filter will be created
            if (numberOfColumns == 0)
            {
                return null;
            }

            AutoFilter filter = new AutoFilter
            {
                //e.g. A1:D1
                Reference = string.Format("{0}1:{1}1", GetHeaderCellColumnString(1), GetHeaderCellColumnString(numberOfColumns))
            };
            return filter;
        }

        public virtual string GetHeaderCellColumnString(int columnIndex)
        {
            return GetCellColumnInExcelStyle(columnIndex);
        }


        //Generate excel style of cell index number, e.g. A5, B6, AC123, BD5448
        private string GetCellColumnInExcelStyle(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            const int AtoZ = 26;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % AtoZ;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = ((dividend - modulo) / AtoZ);
            }

            return columnName;
        }
    }
}
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IAutoFilterProvider
    {
        AutoFilter CreateAutoFilter(SheetData sheetData);
        string GetHeaderCellColumnString(int columnIndex);

    }
}
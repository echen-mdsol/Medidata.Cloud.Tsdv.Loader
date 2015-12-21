using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ICellTypeValueConverterFactory
    {
        void GetCellTypeAndValue(object value, out CellValues cellType, out string cellValue);
        object GetCSharpValue(Cell cell);
    }
}
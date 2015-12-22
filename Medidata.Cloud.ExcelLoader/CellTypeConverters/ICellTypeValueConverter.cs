using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    public interface ICellTypeValueConverter
    {
        CellValues CellType { get; }
        bool TryConvertToCellValue(object csharpValue, out string cellValue);
        bool TryConvertToCSharpValue(string cellValue, out object csharpValue);
    }
}
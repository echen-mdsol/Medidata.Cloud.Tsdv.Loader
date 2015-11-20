using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    public interface ICellTypeValueConverter
    {
        CellValues CellType { get; }
        Type CSharpType { get; }
        string GetCellValue(object csharpValue);
        object GetCSharpValue(string cellValue);
    }
}
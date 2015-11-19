using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    public interface ICellValueConverter
    {
        CellValues CellType { get; }
        Type CSharpType { get; }
        string GetCellValue(object csharpValue);
        object GetCSharpValue(string cellValue);
    }
}
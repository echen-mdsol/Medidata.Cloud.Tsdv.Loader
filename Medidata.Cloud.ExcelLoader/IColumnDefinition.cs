using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IColumnDefinition
    {
        CellValues CellType { get; set; }
        Type PropertyType { get; set; }
        string PropertyName { get; set; }
    }
}
using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public class ColumnDefinition : IColumnDefinition
    {
        public CellValues CellType { get; set; }
        public Type PropertyType { get; set; }
        public string PropertyName { get; set; }
    }
}
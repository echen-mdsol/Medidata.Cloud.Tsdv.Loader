using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IColumnDefinition
    {
        Type PropertyType { get; set; }
        string PropertyName { get; set; }
        string HeaderName { get; set; }
    }
}
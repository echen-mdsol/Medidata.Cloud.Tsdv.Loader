using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder
    {
        Action<IEnumerable<object>, ISheetDefinition, SpreadsheetDocument> BuildSheet { get; set; }
        Func<object, ISheetDefinition, Row> BuildRow { get; set; }
    }
}
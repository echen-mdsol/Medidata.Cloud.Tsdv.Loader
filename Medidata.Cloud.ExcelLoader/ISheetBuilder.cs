using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder
    {
        Action<IEnumerable<SheetModel>, ISheetDefinition, SpreadsheetDocument> BuildSheet { get; set; }
        Func<SheetModel, ISheetDefinition, Row> BuildRow { get; set; }
    }
}
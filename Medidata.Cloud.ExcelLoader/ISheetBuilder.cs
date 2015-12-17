using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder
    {
        Action<IEnumerable<SheetDefinitionModelBase>, ISheetDefinition, SpreadsheetDocument> BuildSheet { get; set; }
        Func<SheetDefinitionModelBase, ISheetDefinition, Row> BuildRow { get; set; }
    }
}
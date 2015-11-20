using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    internal interface ISheetBuilder : IList<object>
    {
        string SheetName { get; }
        bool HasHeaderRow { get; }
        string[] ColumnNames { get; }
        void AttachTo(SpreadsheetDocument doc);
    }
}
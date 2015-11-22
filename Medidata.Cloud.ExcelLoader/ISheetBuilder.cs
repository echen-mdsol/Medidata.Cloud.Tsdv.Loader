using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilder : IList
    {
        string SheetName { get; }
        bool HasHeaderRow { get; }
        string[] ColumnNames { get; }
        void AttachTo(SpreadsheetDocument doc);
    }
}
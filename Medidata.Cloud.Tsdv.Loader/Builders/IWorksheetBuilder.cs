using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface IWorksheetBuilder: IList<object>
    {
        string[] ColumnNames { get; set; }
        void AppendWorksheet(SpreadsheetDocument doc, bool hasHeaderRow, string sheetName);
    }
}
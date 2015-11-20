using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface ISpreadsheetBuilder
    {
        IList<T> EnsureWorksheet<T>(string sheetName, bool hasHeaderRow = true, string[] columnNames = null)
            where T : class;

        SpreadsheetDocument Save(Stream outStream);
    }
}
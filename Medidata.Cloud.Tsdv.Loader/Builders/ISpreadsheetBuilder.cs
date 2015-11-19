using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface ISpreadsheetBuilder
    {
        IList<T> EnsureWorksheet<T>(string sheetName, string [] columnNames = null) where T : class;
        SpreadsheetDocument Save(Stream outStream);
    }
}
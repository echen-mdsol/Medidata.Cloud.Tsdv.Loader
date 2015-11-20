using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface ISpreadsheetBuilder
    {
        IList<object> EnsureWorksheet<T>(string sheetName, string[] columnNames) where T : class;
        IList<object> EnsureWorksheet<T>(string sheetName, bool hasHeaderRow) where T : class;
        void Save(Stream outStream);
    }
}
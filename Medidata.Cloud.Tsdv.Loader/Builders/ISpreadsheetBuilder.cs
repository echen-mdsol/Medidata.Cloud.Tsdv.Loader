using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public interface ISpreadsheetBuilder
    {
        IList<object> EnsureWorksheet<T>(string sheetName, bool hasHeaderRow = true, string[] columnNames = null)
            where T : class;

        void Save(Stream outStream);
    }
}
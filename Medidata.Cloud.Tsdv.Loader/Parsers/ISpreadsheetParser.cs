using System;
using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public interface ISpreadsheetParser : IDisposable
    {
        IEnumerable<T> RetrieveObjectsFromSheet<T>(string sheetName) where T : class;
        void Load(Stream stream, bool hasHeaderRow = true);
    }
}
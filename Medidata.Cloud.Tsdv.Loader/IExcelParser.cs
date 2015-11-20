using System;
using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IExcelParser : IDisposable
    {
        IEnumerable<T> GetObjects<T>(string sheetName, bool hasHeaderRow = true) where T : class;
        void Load(Stream stream);
    }
}
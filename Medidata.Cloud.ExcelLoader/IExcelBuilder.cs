using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelBuilder
    {
        IList<object> AddSheet<T>(string sheetName, string[] columnNames) where T : class;
        IList<object> AddSheet<T>(string sheetName, bool hasHeaderRow) where T : class;
        IList<object> GetSheet(string sheetName);
        void Save(Stream outStream);
    }
}
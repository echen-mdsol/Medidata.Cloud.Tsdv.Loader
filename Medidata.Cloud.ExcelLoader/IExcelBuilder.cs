using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelBuilder
    {
        IList<T> AddSheet<T>(string sheetName, string headerStyleName, string textStyleName, string[] columnNames) where T : class;
        IList<T> AddSheet<T>(string sheetName, string headerStyleName, string textStyleName, bool hasHeaderRow = true) where T : class;
        IList<T> GetSheet<T>(string sheetName) where T : class;
        void Save(Stream outStream);
    }
}
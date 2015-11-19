using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    public interface IWorksheetParser<out T> where T : class
    {
        IEnumerable<T> GetObjects();
        void Load(Worksheet worksheet);
    }
}
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Parsers
{
    internal interface ISheetParser<out T> where T : class
    {
        bool HasHeaderRow { get; }
        IEnumerable<T> GetObjects();
        void Load(Worksheet worksheet);
    }
}
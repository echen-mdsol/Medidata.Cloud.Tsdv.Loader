using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IOpenXmlWorksheetBuilder<T> where T: class
    {
        void AddObject(T target);
        Worksheet ToOpenXmlWorksheet();
    }
}
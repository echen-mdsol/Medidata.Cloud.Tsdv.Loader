using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader
{
    public interface IWorksheetBuilder
    {
        Sheet ToWorksheet(string name);
    }
}
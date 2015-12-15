using System;
using System.Collections.Generic;
using System.IO;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelLoader : IDisposable
    {
        ISheetBroker Sheet(ISheetDefinition sheetDefinition, params ISheetDecorator[] decorators);
        void Save(Stream outStream);
        void Load(Stream stream);
    }
}
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using Medidata.Cloud.ExcelLoader.Helpers;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public interface IExcelLoader
    {
        void Save(Stream outStream);
        void Load(Stream source);
        ISheetInfo<T> Sheet<T>() where T: SheetModel;
    }

    public interface ISheetInfo<T>  where T: SheetModel
    {
        ISheetDefinition Definition { get; }
        IList<T> Data { get; }
    }


}
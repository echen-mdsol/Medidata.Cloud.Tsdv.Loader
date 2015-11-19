using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.CellValueConverters
{
    internal abstract class CellValueBaseConverter<T> : ICellValueConverter
    {
        public CellValues CellType { get; private set; }
        public Type CSharpType { get; private set; }

        protected CellValueBaseConverter(CellValues cellType)
        {
            CellType = cellType;
            CSharpType = typeof (T);
        }

        protected abstract string GetCellValueImpl(T csharpValue);

        public string GetCellValue(object csharpValue)
        {
            return GetCellValueImpl((T)csharpValue);
        }

        protected abstract T GetCSharpValueImpl(string cellValue);

        public object GetCSharpValue(string cellValue)
        {
            return GetCSharpValueImpl(cellValue);
        }
    }
}
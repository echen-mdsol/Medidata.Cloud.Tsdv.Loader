using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Medidata.Cloud.ExcelLoader.CellTypeConverters
{
    internal abstract class CellTypeValueBaseConverter<T> : ICellTypeValueConverter
    {
        protected CellTypeValueBaseConverter(CellValues cellType)
        {
            CellType = cellType;
            CSharpType = typeof (T);
        }

        public CellValues CellType { get; private set; }
        public Type CSharpType { get; private set; }

        public string GetCellValue(object csharpValue)
        {
            return csharpValue == null ? null : GetCellValueImpl((T)csharpValue);
        }

        public object GetCSharpValue(string cellValue)
        {
            return GetCSharpValueImpl(cellValue);
        }
        protected bool IsNumeric(string value){
            var isNumericRegex = new Regex("^(" +
                                            /*Hex*/ @"0x[0-9a-f]+" + "|" +
                                            /*Bin*/ @"0b[01]+" + "|" +
                                            /*Oct*/ @"0[0-7]*" + "|" +
                                            /*Dec*/ @"((?!0)|[-+]|(?=0+\.))(\d*\.)?\d+(e\d+)?" +")$");
           return isNumericRegex.IsMatch(value);
        }
        //protected abstract T GetCellValueImpl(T csharpValue);
        protected virtual string GetCellValueImpl(T csharpValue)
        {
            return csharpValue.ToString();
        }

        protected abstract T GetCSharpValueImpl(string cellValue);
    }
}
using System;

namespace Medidata.Cloud.ExcelLoader.SheetDefinitions
{
    [AttributeUsage(AttributeTargets.Class, Inherited = false)]
    public class SheetNameAttribute : Attribute
    {
        public SheetNameAttribute(string sheetName)
        {
            SheetName = sheetName;
        }

        public string SheetName { get; set; }
    }
}
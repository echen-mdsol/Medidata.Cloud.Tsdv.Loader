using System;

namespace Medidata.Cloud.ExcelLoader.SheetDefinitions
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnHeaderNameAttribute : Attribute
    {
        public ColumnHeaderNameAttribute(string header)
        {
            Header = header;
        }

        public string Header { get; set; }
    }
}
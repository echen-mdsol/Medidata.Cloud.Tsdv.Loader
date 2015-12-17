using System;

namespace Medidata.Cloud.ExcelLoader
{
    public class ColumnDefinition : IColumnDefinition
    {
        public static dynamic New
        {
            get { return new ColumnBuilder(); }
        }

        public Type PropertyType { get; set; }
        public string PropertyName { get; set; }
        public string HeaderName { get; set; }
    }
}
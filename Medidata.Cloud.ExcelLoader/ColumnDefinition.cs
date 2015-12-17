using System;
using System.Dynamic;

namespace Medidata.Cloud.ExcelLoader
{
    public class ColumnDefinition : IColumnDefinition
    {
        public Type PropertyType { get; set; }
        public string PropertyName { get; set; }
        public string HeaderName { get; set; }

        public static dynamic New
        {
            get { return new ColumnBuilder(); }
        }
    }

    public class ColumnBuilder : DynamicObject
    {
        public override bool TryInvoke(InvokeBinder binder, object[] args, out object result)
        {

            result = new ColumnDefinition();
            return true;
        }
    }
}
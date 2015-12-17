using System.Dynamic;

namespace Medidata.Cloud.ExcelLoader
{
    public class ColumnBuilder : DynamicObject
    {
        public override bool TryInvoke(InvokeBinder binder, object[] args, out object result)
        {
            result = new ColumnDefinition();
            return true;
        }
    }
}
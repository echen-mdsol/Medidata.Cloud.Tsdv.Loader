using System.Collections.Generic;
using System.Dynamic;

namespace Medidata.Cloud.ExcelLoader.SheetDefinitions
{
    public abstract class SheetDefinitionModelBase : DynamicObject
    {
        protected SheetDefinitionModelBase()
        {
            ExtraProperties = new ExpandoObject();
        }

        [ColumnIngored]
        public IDictionary<string, object> ExtraProperties { get; private set; }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return ExtraProperties.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            var propName = binder.Name;
            if (ExtraProperties.ContainsKey(propName))
            {
                ExtraProperties[propName] = value;
            }
            else
            {
                ExtraProperties.Add(propName, value);
            }
            return true;
        }
    }
}
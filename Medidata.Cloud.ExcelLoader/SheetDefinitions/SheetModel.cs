using System.Collections.Generic;
using System.Dynamic;

namespace Medidata.Cloud.ExcelLoader.SheetDefinitions
{
    public abstract class SheetModel : DynamicObject
    {
        private readonly IDictionary<string, object> _extraProperties = new ExpandoObject();

        public IDictionary<string, object> GetExtraProperties()
        {
            return _extraProperties;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            return _extraProperties.TryGetValue(binder.Name, out result);
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            var propName = binder.Name;
            if (_extraProperties.ContainsKey(propName))
            {
                _extraProperties[propName] = value;
            }
            else
            {
                _extraProperties.Add(propName, value);
            }
            return true;
        }
    }
}
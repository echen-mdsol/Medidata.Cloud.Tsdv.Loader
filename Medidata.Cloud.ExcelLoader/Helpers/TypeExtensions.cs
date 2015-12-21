using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    internal static class TypeExtensions
    {
        public static IEnumerable<PropertyDescriptor> GetPropertyDescriptors(this Type type)
        {
            return TypeDescriptor.GetProperties(type).OfType<PropertyDescriptor>();
        }

        public static object GetPropertyValue(this object target, string propertyName)
        {
            var prop = GetPropertyDescriptors(target.GetType()).FirstOrDefault(x => x.Name == propertyName);
            return prop == null ? null : prop.GetValue(target);
        }
    }
}
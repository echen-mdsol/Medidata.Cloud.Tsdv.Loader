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

        public static PropertyDescriptor GetPropertyDescriptor(this Type type, string propertyName)
        {
            return GetPropertyDescriptors(type).FirstOrDefault(x => x.Name == propertyName);
        }

        public static object GetPropertyValue(this object target, string propName)
        {
            var prop = GetPropertyDescriptor(target.GetType(), propName);
            return prop == null ? null : prop.GetValue(target);
        }
    }
}
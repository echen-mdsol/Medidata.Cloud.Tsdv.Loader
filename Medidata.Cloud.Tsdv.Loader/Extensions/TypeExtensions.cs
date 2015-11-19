using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Medidata.Cloud.Tsdv.Loader.Extensions
{
    internal static class TypeExtensions
    {
        public static IEnumerable<PropertyDescriptor> GetPropertyDescriptors(this Type type)
        {
            return TypeDescriptor.GetProperties(type).OfType<PropertyDescriptor>();
        } 
    }
}

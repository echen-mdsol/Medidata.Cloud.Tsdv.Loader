using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class TypeExtensions
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
            try
            {
                var property = target.GetType().GetPropertyDescriptor(propName);
                return property.GetValue(target);
            }
            catch
            {
                return null;
            }
        }

        public static ISheetDefinition GetSheetDefinitionFromType(this Type target, string sheetName,
            IEnumerable<string> headers = null)
        {
            var props = target.GetPropertyDescriptors();
            var headerList = (headers ?? Enumerable.Empty<string>()).ToList();
            var sheetDefinition = new SheetDefinition
            {
                Name = sheetName,
                ColumnDefinitions =
                    props.Select((x, i) => PropertyToColumnDefinition(x, i < headerList.Count ? headerList[i] : null))
            };
            return sheetDefinition;
        }

        private static IColumnDefinition PropertyToColumnDefinition(PropertyDescriptor property,
            string headerName = null)
        {
            return new ColumnDefinition
            {
                PropertyType = property.PropertyType,
                PropertyName = property.Name,
                HeaderName = headerName ?? property.Name
            };
        }
    }
}
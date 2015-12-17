using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.Linq;
using System.Runtime.Remoting.Contexts;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class TypeToSheetDefinitionExtensions
    {
        public static ISheetDefinition GetSheetDefinitionFromType(this Type target, string sheetName, IEnumerable<string> headers = null)
        {
            var props = target.GetPropertyDescriptors();
            var headerList = (headers ?? Enumerable.Empty<string>()).ToList();
            var sheetDefinition = new SheetDefinition
            {
                Name = sheetName,
                ColumnDefinitions = props.Select((x, i) => PropertyToColumnDefinition(x, i < headerList.Count ? headerList[i] : null)).ToList()
            };

            return sheetDefinition;
        }

        public static bool IsDynamicPropertyObject(Type type)
        {
            return type.GetInterfaces().Any(x => x == typeof(IDynamicProperty));
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
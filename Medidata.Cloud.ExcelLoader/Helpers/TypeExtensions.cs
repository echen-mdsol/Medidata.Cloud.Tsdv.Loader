using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

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

        public static ISheetDefinition ToSheetDefinition(this Type type, string sheetName, ICellTypeValueConverterFactory converterFactory = null)
        {
            var props = GetPropertyDescriptors(type);
            var sheetDefinition = new SheetDefinition
            {
                Name = sheetName,
                ColumnDefinitions = props.Select(x => PropertyToColumnDefinition(x, converterFactory))
            };
            return sheetDefinition;
        }

        private static IColumnDefinition PropertyToColumnDefinition(PropertyDescriptor property, ICellTypeValueConverterFactory converterFactory)
        {
            var factory = converterFactory ?? new CellTypeValueConverterFactory();
            var converter = factory.Produce(property.PropertyType);
            return new ColumnDefinition
            {
                PropertyType = property.PropertyType,
                PropertyName = property.Name,
                CellType = converter.CellType
            };
        }
    }
}
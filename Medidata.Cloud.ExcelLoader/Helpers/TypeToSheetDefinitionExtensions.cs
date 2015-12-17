using System;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class TypeToSheetDefinitionExtensions
    {
        public static ISheetDefinition ToSheetDefinition(this Type type)
        {
            var sheetNameAtt = type.GetCustomAttributes(false).OfType<SheetNameAttribute>().SingleOrDefault();
            if (sheetNameAtt == null)
                throw new Exception("Model type must have SheetNameAttribute");

            var props = type.GetPropertyDescriptors();
            var colDefinitions = from prop in props
                                 let headerAtt = prop.Attributes.OfType<ColumnHeaderNameAttribute>().SingleOrDefault()
                                 let ignoreAtt = prop.Attributes.OfType<ColumnIngoredAttribute>().SingleOrDefault()
                                 where ignoreAtt == null
                                 select new ColumnDefinition
                                        {
                                            PropertyName = prop.Name,
                                            PropertyType = prop.PropertyType,
                                            Header = headerAtt != null ? headerAtt.Header : prop.Name
                                        };

            var sheetDefinition = new SheetDefinition
                                  {
                                      Name = sheetNameAtt.SheetName,
                                      ColumnDefinitions = colDefinitions
                                  };
            return sheetDefinition;
        }

        public static bool IsDynamicPropertyObject(Type type)
        {
            return type.GetInterfaces().Any(x => x == typeof(IDynamicProperty));
        }

        //        private static IColumnDefinition PropertyToColumnDefinition(PropertyDescriptor property,

        //                                                                    string headerName = null)
        //        {
        //            return new ColumnDefinition
        //                   {
        //                       PropertyType = property.PropertyType,
        //                       PropertyName = property.Name,
        //                       HeaderName = headerName ?? property.Name
        //                   };
        //        }
    }
}
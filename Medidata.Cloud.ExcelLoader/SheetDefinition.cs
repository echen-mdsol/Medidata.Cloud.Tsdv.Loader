using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Medidata.Cloud.ExcelLoader.Helpers;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public class SheetDefinition : ISheetDefinition
    {
        internal SheetDefinition() {}

        public string Name { get; set; }
        public bool AcceptExtraProperties { get; set; }
        public IEnumerable<IColumnDefinition> ColumnDefinitions { get; set; }

        public ISheetDefinition AddColumn(IColumnDefinition columnDefinition)
        {
            if (ColumnDefinitions.Any(x => x.PropertyName == columnDefinition.PropertyName))
            {
                throw new Exception("Duplicate column binding to propety " + columnDefinition.PropertyName);
            }
            ColumnDefinitions = ColumnDefinitions.Concat(new[] {columnDefinition});
            return this;
        }

        public static ISheetDefinition Define<T>()
        {
            var type = typeof(T);
            var sheetNameAtt = type.GetCustomAttributes(false).OfType<SheetNameAttribute>().SingleOrDefault();
            if (sheetNameAtt == null)
                throw new Exception("Model type must have SheetNameAttribute");

            var props = type.GetPropertyDescriptors();
            var colDefinitions = from prop in props
                                 let headerAtt = prop.Attributes.OfType<ColumnHeaderNameAttribute>().SingleOrDefault()
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

        private IColumnDefinition PropToConDef(PropertyDescriptor property, string headerName = null)
        {
            return new ColumnDefinition
                   {
                       PropertyType = property.PropertyType,
                       PropertyName = property.Name,
                       Header = headerName ?? property.Name
                   };
        }
    }
}
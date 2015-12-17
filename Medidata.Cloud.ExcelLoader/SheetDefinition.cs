using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Medidata.Cloud.ExcelLoader.Helpers;

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
            return typeof(T).ToSheetDefinition();
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
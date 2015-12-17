using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.Linq;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader
{
    public class SheetDefinition : ISheetDefinition
    {
        public string Name { get; private set; }
        public bool AcceptExtraProperties { get; set; }
        public IEnumerable<IColumnDefinition> ColumnDefinitions { get; set; }

        public SheetDefinition(string sheetName)
        {
            Name = sheetName;
            ColumnDefinitions = Enumerable.Empty<IColumnDefinition>();
        }

        public ISheetDefinition DefineColumns<T>(params string[] headers)
        {
            var props = typeof(T).GetPropertyDescriptors();
            ColumnDefinitions = props.Select((x, i) => PropToConDef(x, i < headers.Length ? headers[i] : null));
            return this;
        }

        public ISheetDefinition AddColumn(IColumnDefinition columnDefinition)
        {
            if (ColumnDefinitions.Any(x => x.PropertyName == columnDefinition.PropertyName))
            {
                throw new Exception("Duplicate column binding to propety " + columnDefinition.PropertyName);
            }
            ColumnDefinitions = ColumnDefinitions.Concat(new [] { columnDefinition});
            return this;
        }

        private IColumnDefinition PropToConDef(PropertyDescriptor property, string headerName = null)
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
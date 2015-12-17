using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class ListExtensions
    {
        public static IList<T> AddRange<T>(this IList<T> target, IEnumerable<T> items)
        {
            if (items == null) return target;
            foreach (var item in items)
            {
                target.Add(item);
            }
            return target;
        }

        public static IList<T> Add<T>(this IList<T> target, params T[] items)
        {
            return AddRange(target, items);
        }

        public static IEnumerable<T> OfSheetModel<T>(this IEnumerable<ExpandoObject> sourceList) where T : SheetModel
        {
            var type = typeof(T);
            var ownProps =
                type.GetPropertyDescriptors().Where(p => !p.Attributes.OfType<ColumnIngoredAttribute>().Any()).ToList();
            foreach (var item in sourceList)
            {
                var model = Activator.CreateInstance<T>();
                var extraPropsOwner = (SheetModel) model;

                foreach (var sourceKvp in item)
                {
                    var sourcePropName = sourceKvp.Key;
                    var sourcePropValue = sourceKvp.Value;
                    var ownProp = ownProps.FirstOrDefault(p => p.Name == sourcePropName);
                    if (ownProp != null)
                    {
                        // Own property
                        ownProp.SetValue(model, sourcePropValue);
                    }
                    else
                    {
                        // Extra property
                        extraPropsOwner.ExtraProperties.Add(sourcePropName, sourcePropValue);
                    }
                }

                yield return model;
            }
        }
    }
}
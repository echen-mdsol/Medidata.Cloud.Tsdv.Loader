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

        internal static IEnumerable<T> CastToSheetModel<T>(this IEnumerable<ExpandoObject> target) where T : SheetModel
        {
            var type = typeof(T);
            var ownProps = type.GetPropertyDescriptors().ToList();
            foreach (var item in target)
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
                        extraPropsOwner.GetExtraProperties().Add(sourcePropName, sourcePropValue);
                    }
                }

                yield return model;
            }
        }
    }
}
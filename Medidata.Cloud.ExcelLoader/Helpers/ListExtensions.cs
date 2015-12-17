using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using ImpromptuInterface;

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

        private static ExpandoObject ConvertToExpando<T>(object target)
        {
            if (target == null || target is ExpandoObject) return (ExpandoObject)target;
            var props = typeof(T).GetPropertyDescriptors();
            var expando = new ExpandoObject();
            IDictionary<string, object> dic = expando;
            foreach (var prop in props)
            {
                dic.Add(prop.Name, target.GetPropertyValue(prop.Name));
            }
            return expando;
        }

        public static IList<T> Add<T>(this IList<T> target, params T[] items)
        {
            return AddRange(target, items);
        }
//
//        public static IList<T> AddLike<T>(this IList<T> target, params object[] items) where T : class
//        {
//            return AddRange(target, items.Select(x => x.ActLike<T>()));
//        }
//
//        public static IEnumerable<T> ActAs<T>(this IEnumerable<ExpandoObject> target) where T : class
//        {
//            return target.Select(x => x.ActLike<T>());
//        }
    }
}
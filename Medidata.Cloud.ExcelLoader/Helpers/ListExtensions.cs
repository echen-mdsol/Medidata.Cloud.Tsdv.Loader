using System.Collections.Generic;

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

        //        public static IList<T> AddLike<T>(this IList<T> target, params object[] items) where T : class

        //
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
using System.Collections.Generic;
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

        public static IList<T> Add<T>(this IList<T> target, params T[] items)
        {
            return AddRange(target, items);
        }

        public static IList<T> AddAnonymous<T>(this IList<T> target, params object[] items) where T : class
        {
            return AddRange(target, items.Select(x => x.ActLike<T>()));
        }
    }
}
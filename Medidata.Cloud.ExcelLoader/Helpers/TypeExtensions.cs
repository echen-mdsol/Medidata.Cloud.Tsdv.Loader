using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.Linq;
using ImpromptuInterface;

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

        public static object GetPropertyValue(this object target, string propName)
        {
            try
            {
                return Impromptu.InvokeGet(target, propName);
            }
            catch
            {
                return null;
            }
        }
    }


    public interface IExtraProperty
    {
        ExpandoObject ExtraProperties { get; }
    }

    public static class ObjectExtensions
    {
        public static IList<T> AddSimilarShape<T>(this IList<T> target, params object[] items) where T: class
        {
            var newItems = items.Select(ActAs<T>);
            target.AddRange(newItems);
            return target;
        } 

        internal static T ActAs<T>(this object target) where T : class
        {
            if (target == null) return null;
            return IsDynamicPropertyObject<T>() ? ActWithExtraProperties<T>(target) : target.ActLike<T>();
        }

        private static T ActWithExtraProperties<T>(object target) where T : class
        {
            if(!typeof(IExtraProperty).IsAssignableFrom(typeof(T)))
                throw new Exception("T must be of type " + typeof(IExtraProperty));

            dynamic expando = new ExpandoObject();
            expando.ExtraProperties = new ExpandoObject();

            IDictionary<string, object> expandoDic = expando;
            IDictionary<string, object> extraPropPropertyDic = expando.ExtraProperties;

            var typeProps = typeof(T).GetPropertyDescriptors().ToList();
            foreach (var typeProp in typeProps)
            {
                var propValue = target.GetPropertyValue(typeProp.Name);
                expandoDic.Add(typeProp.Name, propValue);
            }
            var props = target.GetType().GetPropertyDescriptors();
            var extraProps = props.Except(typeProps);

            foreach (var extraProp in extraProps)
            {
                extraPropPropertyDic.Add(extraProp.Name, target.GetPropertyValue(extraProp.Name));
            }

            T actor = Impromptu.ActLike(expando);
            return actor;
        }

        private static bool IsDynamicPropertyObject<T>()
        {
            return typeof(T).GetInterfaces().Any(x => x == typeof(IExtraProperty));
        }
    }
}
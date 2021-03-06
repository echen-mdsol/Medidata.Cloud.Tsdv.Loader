namespace Medidata.Cloud.ExcelLoader.SheetDefinitions
{
    public static class SheetModelExtensions
    {
        public static T AddProperty<T>(this T target, string propName, object value) where T : SheetModel
        {
            target.GetExtraProperties().Add(propName, value);
            return target;
        }
    }
}
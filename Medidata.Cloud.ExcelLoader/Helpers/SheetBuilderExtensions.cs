namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class SheetBuilderExtensions
    {
        public static ISheetBuilder Decorate(this ISheetBuilder target, ISheetBuilderDecorator decorator)
        {
            return decorator.Decorate(target);
        }
    }
}
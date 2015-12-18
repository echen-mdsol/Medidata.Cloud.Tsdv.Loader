namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilderDecorator
    {
        ISheetBuilder Decorate(ISheetBuilder target);
    }
}
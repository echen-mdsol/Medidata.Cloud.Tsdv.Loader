namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetBuilderFactory
    {
        ISheetBuilder Create<T>() where T : class;
    }
}
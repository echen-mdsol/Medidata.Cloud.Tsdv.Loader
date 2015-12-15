namespace Medidata.Cloud.ExcelLoader.SheetDecorators
{
    public class AutoFitWidthSheetDecorator : ISheetBuilderDecorator
    {
        public ISheetBuilder Decorate(ISheetBuilder target)
        {
            var originalFunc = target.BuildSheet;
            // TODO: AutoFit feature is pending.

            return target;
        }
    }
}
using System.Dynamic;

namespace Medidata.Cloud.ExcelLoader.SheetDecorators
{
    public class ExpandoObjectSheetDecorator : ISheetBuilderDecorator
    {
        public ISheetBuilder Decorate(ISheetBuilder target)
        {
            var originalFunc = target.BuildRow;
            target.BuildRow = (model, sheetDefinition) =>
            {
                if (model is ExpandoObject) {}

                return originalFunc(model, sheetDefinition);
            };

            return target;
        }
    }
}
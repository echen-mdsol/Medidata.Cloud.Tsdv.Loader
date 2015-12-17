using System.Dynamic;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.Helpers;

namespace Medidata.Cloud.ExcelLoader.SheetDecorators
{
    public class TextStyleSheetDecorator : ISheetBuilderDecorator
    {
        private readonly string _styleName;

        public TextStyleSheetDecorator(string styleName)
        {
            _styleName = styleName;
        }

        public ISheetBuilder Decorate(ISheetBuilder target)
        {
            var originalFunc = target.BuildSheet;
            target.BuildSheet = (models, sheetDefinition, doc) =>
            {
                originalFunc(models, sheetDefinition, doc);
                var styleIndex = doc.GetStyleIndex(_styleName);
                var sheetData = doc.GetSheetDataByName(sheetDefinition.Name);
                var cells = sheetData.Descendants<Cell>();
                foreach (var cell in cells)
                {
                    cell.StyleIndex = styleIndex;
                }
            };

            return target;
        }
    }

    public class ExpandoObjectSheetDecorator : ISheetBuilderDecorator
    {
        public ISheetBuilder Decorate(ISheetBuilder target)
        {
             var originalFunc = target.BuildRow;
            target.BuildRow = (model, sheetDefinition) =>
            {
                if (model is ExpandoObject)
                {
                    
                }

                return originalFunc(model, sheetDefinition);
            };

            return target;
        }
    }
}
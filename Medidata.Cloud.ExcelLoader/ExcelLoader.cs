using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Medidata.Cloud.ExcelLoader
{
    public class ExcelLoader : IExcelLoader
    {
        private ExcelPackage _package;
        private readonly IDictionary<string, ISheetDecorator[]> _decoratorDic = new Dictionary<string, ISheetDecorator[]>(); 

        public ExcelLoader()
        {
            _package = new ExcelPackage();
        }

        public ISheetBroker Sheet(ISheetDefinition sheetDefinition, params ISheetDecorator[] decorators)
        {
            var sheetName = sheetDefinition.Name;
            var sheet = _package.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName) ??
                        _package.Workbook.Worksheets.Add(sheetName);

            var sheetBroker = new SheetBroker();
            sheetBroker.Worksheet = sheet;
            sheetBroker.SheetDefinition = sheetDefinition;

            _decoratorDic.Add(sheetName, decorators);

            return sheetBroker;
        }

        public void Save(Stream outStream)
        {
            foreach (var kvp in _decoratorDic)
            {
                var sheet = _package.Workbook.Worksheets[kvp.Key];
                foreach (var decorator in kvp.Value)
                {
                    decorator.Decorate(sheet);
                }
            }
            _package.SaveAs(outStream);
        }

        public void Load(Stream stream)
        {
            Dispose();
            _package = new ExcelPackage();
            _package.Load(stream);
        }

        public void Dispose()
        {
            if (_package != null)
            {
                _package.Dispose();
                _package = null;
            }
        }
    }
}
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader
{

    public class ExcelBuilder : IExcelBuilder
    {
        private class SheetModels : List<object>
        {
            public ISheetDefinition SheetDefinition { get; set; }
            public ISheetBuilder SheetBuilder { get; set; }
        }

        private readonly IDictionary<string, SheetModels> _modelDic = new Dictionary<string, SheetModels>();

        public IList<object> DefineSheet(ISheetDefinition sheetDefinition, ISheetBuilder sheetBuilder)
        {
            var sheetName = sheetDefinition.Name;
            _modelDic.Add(sheetName, new SheetModels { SheetDefinition = sheetDefinition, SheetBuilder = sheetBuilder});
            return _modelDic[sheetName];
        }

        public virtual void Save(Stream outStream)
        {
            using (var doc = CreateDocument(outStream))
            {
                WorkbookPart workbookPart;
                if (doc.WorkbookPart == null)
                {
                    workbookPart = doc.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                }
                else
                {
                    workbookPart = doc.WorkbookPart;
                }

                foreach (var key in _modelDic.Keys)
                {
                    var info = _modelDic[key];
                    info.SheetBuilder.BuildSheet(info, info.SheetDefinition, doc);
                }

                workbookPart.Workbook.Save();
            }
        }
        protected virtual SpreadsheetDocument CreateDocument(Stream outStream)
        {
            return SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook);
        }
    }
}
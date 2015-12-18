using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Medidata.Cloud.ExcelLoader.SheetDefinitions;

namespace Medidata.Cloud.ExcelLoader
{
    public class ExcelBuilder : IExcelBuilder
    {
        private readonly IDictionary<string, SheetModels> _modelDic = new Dictionary<string, SheetModels>();

        public void AddSheet(ISheetDefinition sheetDefinition, IEnumerable<SheetModel> models,
                             ISheetBuilder sheetBuilder)
        {
            if (sheetDefinition == null) throw new ArgumentNullException("sheetDefinition");
            if (models == null) throw new ArgumentNullException("models");
            if (sheetBuilder == null) throw new ArgumentNullException("sheetBuilder");
            var sheetName = sheetDefinition.Name;
            var sheetModels = new SheetModels {SheetDefinition = sheetDefinition, SheetBuilder = sheetBuilder};
            sheetModels.AddRange(models);
            _modelDic.Add(sheetName, sheetModels);
        }

        public virtual void Save(Stream outStream)
        {
            if (outStream == null) throw new ArgumentNullException("outStream");
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
            if (outStream == null) throw new ArgumentNullException("outStream");
            return SpreadsheetDocument.Create(outStream, SpreadsheetDocumentType.Workbook);
        }

        private class SheetModels : List<SheetModel>
        {
            public ISheetDefinition SheetDefinition { get; set; }
            public ISheetBuilder SheetBuilder { get; set; }
        }
    }
}
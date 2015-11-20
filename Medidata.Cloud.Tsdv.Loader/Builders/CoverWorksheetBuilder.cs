using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class CoverWorksheetBuilder : List<object>, IWorksheetBuilder
    {
        private static Worksheet _coverSheet;
        private static readonly object CoverSheetLock = new object();
        public string[] ColumnNames { get; set; }

        public void AppendWorksheet(SpreadsheetDocument doc, bool hasHeaderRow, string sheetName)
        {
            var worksheet = CreateWorksheet();
            WorksheetBuilderHelper.AppendWorksheet(doc, worksheet, sheetName);
        }

        public Worksheet CreateWorksheet()
        {
            if (_coverSheet != null) return _coverSheet;
            lock (CoverSheetLock)
            {
                if (_coverSheet != null) return _coverSheet;
                var sheetBytes = Resource.CoverSheet;
                using (var ms = new MemoryStream())
                {
                    ms.Write(sheetBytes, 0, sheetBytes.Length);
                    using (var ss = SpreadsheetDocument.Open(ms, false))
                    {
                        var coverSheetId = ss.WorkbookPart.Workbook.Descendants<Sheet>().First().Id;
                        var worksheetPart = (WorksheetPart) ss.WorkbookPart.GetPartById(coverSheetId);
                        _coverSheet = worksheetPart.Worksheet;
                    }
                }
            }

            return _coverSheet;
        }
    }
}
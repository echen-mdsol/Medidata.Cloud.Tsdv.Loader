using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.Tsdv.Loader.Builders
{
    public class CoverWorksheetBuilder : IWorksheetBuilder
    {
        public string[] ColumnNames { get; set; }

        private static Sheet _coverSheet;
        private static readonly object CoverSheetLock = new object();
        public Sheet ToWorksheet(string sheetName)
        {
            if (_coverSheet == null)
            {
                lock (CoverSheetLock)
                {
                    if (_coverSheet == null)
                    {

                        var sheetBytes = Resource.CoverSheet;
                        using (var ms = new MemoryStream())
                        {
                            ms.Write(sheetBytes, 0, sheetBytes.Length);
                            using (var ss = SpreadsheetDocument.Open(ms, false))
                            {
                                _coverSheet = ss.WorkbookPart.Workbook.Sheets.Elements<Sheet>().First();
                                _coverSheet.Name = sheetName;
                            }
                        }
                    }
                }
            }

            return _coverSheet;
        }
    }
}
using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    internal class CellStyleProvider
    {
        public uint GetStyleIndex(SpreadsheetDocument doc, string styleName)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            if (styleName == null) throw new ArgumentNullException("styleName");

            var styleSheet = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var cellFormat = FindCellStyleFormat(styleSheet, styleName);

            uint index;
            if (TryGetCellFormatIndex(styleSheet, cellFormat, out index))
            {
                return index;
            }

            index = AddNewCellFormat(styleSheet, cellFormat);
            return index;
        }

        private bool TryGetCellFormatIndex(Stylesheet styleSheet, CellFormat cellFormat, out uint index)
        {
            var cellFormatInfo = styleSheet.CellFormats.Descendants<CellFormat>()
                .Select((c, i) => new {CellFormat = c, Index = (uint) i})
                .FirstOrDefault(c => CellFormatEquals(c.CellFormat, cellFormat));
            var exists = cellFormatInfo != null;
            index = exists ? cellFormatInfo.Index : uint.MaxValue;
            return exists;
        }

        private uint AddNewCellFormat(Stylesheet styleSheet, CellFormat cellFormat)
        {
            var newIndex = (uint) styleSheet.CellFormats.Count();
            var copiedCell = CopyCellFormat(cellFormat);
            styleSheet.CellFormats.AppendChild(copiedCell);
            return newIndex;
        }

        private CellFormat FindCellStyleFormat(Stylesheet styleSheet, string styleName)
        {
            var list = styleSheet.CellStyleFormats.Descendants<CellFormat>().ToList();
            var style = styleSheet.CellStyles.Descendants<CellStyle>().First(x => x.Name == styleName);
            var cellFormat = list[checked((int) (uint) style.FormatId)];
            return cellFormat;
        }

        private bool CellFormatEquals(CellFormat x, CellFormat y)
        {
            return x.FontId == y.FontId &&
                   x.Alignment == y.Alignment &&
                   x.FillId == y.FillId &&
                   x.BorderId == y.BorderId &&
                   x.ApplyNumberFormat == y.ApplyNumberFormat &&
                   x.ApplyBorder == y.ApplyBorder &&
                   x.ApplyAlignment == y.ApplyAlignment &&
                   x.ApplyFill == y.ApplyFill &&
                   x.ApplyFont == y.ApplyFont;
        }

        private CellFormat CopyCellFormat(CellFormat cell)
        {
            return new CellFormat
            {
                FontId = cell.FontId,
                Alignment = cell.Alignment,
                FillId = cell.FillId,
                BorderId = cell.BorderId,
                ApplyNumberFormat = cell.ApplyNumberFormat,
                ApplyBorder = cell.ApplyBorder,
                ApplyAlignment = cell.ApplyAlignment,
                ApplyFill = cell.ApplyFill,
                ApplyFont = cell.ApplyFont
            };
        }
    }
}
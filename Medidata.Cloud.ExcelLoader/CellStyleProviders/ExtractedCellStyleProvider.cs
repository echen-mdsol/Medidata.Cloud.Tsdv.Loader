using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellStyleProviders
{
    /// <summary>
    ///     Extract StyleSheet from an exisiting SpreadSheet
    /// </summary>
    public class ExtractedCellStyleProvider : ICellStyleProvider
    {
        private readonly Dictionary<string, uint> _styleDict = new Dictionary<string, uint>();
        private bool _hasAttachedToDoc;

        public uint GetStyleIndex(string styleName)
        {
            if (!_hasAttachedToDoc) throw new Exception("The style provider hasn't been attached to any document.");
            if (!_styleDict.ContainsKey(styleName))
            {
                throw new ArgumentException("Not supported style name " + styleName);
            }
            return _styleDict[styleName];
        }

        public void AttachTo(SpreadsheetDocument doc)
        {
            var styleSheet = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;

            var newIndex = styleSheet.CellFormats.Count();
            var list = styleSheet.CellStyleFormats.Descendants<CellFormat>().ToList();

            foreach (var style in styleSheet.CellStyles.Descendants<CellStyle>())
            {
                var item = list[checked((int) (uint) style.FormatId)];
                // You cannot append an existing cell format due to the restriction of Openxml sdk, 
                // you need to copy to make a new one. And you need to store the styleindex yourself.
                var copiedCell = CopyCellFormat(item);
                styleSheet.CellFormats.AppendChild(copiedCell);
                _styleDict.Add(style.Name, (uint) newIndex);
                newIndex++;
            }

            _hasAttachedToDoc = true;
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
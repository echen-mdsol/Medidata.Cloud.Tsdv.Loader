using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.CellStyleProviders
{
    /// <summary>
    /// Extract StyleSheet from an exisiting SpreadSheet and use 'Output' and 'Normal' style
    /// </summary>
    public class ExtractedCellStyleProvider: ICellStyleProvider
    {
        private readonly Dictionary<string, uint> _styleDict = new Dictionary<string, uint>();
        public uint GetStyleIndex(string styleName)
        {
            if (!_styleDict.ContainsKey(styleName))
            {
                throw new ArgumentException("styleName");
            }
            return _styleDict[styleName];
        }

        
        public void AttachTo(SpreadsheetDocument doc)
        {
            var styleSheet = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;

            int newIndex = styleSheet.CellFormats.Count();
            var list = styleSheet.CellStyleFormats.Descendants<CellFormat>().ToList();

            foreach (var style in styleSheet.CellStyles.Descendants<CellStyle>())
            {
                var item = list[checked((int)(uint)style.FormatId)];
                //You cannot append an existing cell format due to the restriction of Openxml sdk, you need to copy to make a new one. And you need to store the styleindex yourself.
                styleSheet.CellFormats.AppendChild(CopyCellFormat(item));
                _styleDict.Add(style.Name, (uint)newIndex);
                newIndex++;
            }
        }

        private CellFormat CopyCellFormat(CellFormat cell)
        {
            CellFormat result = new CellFormat();
            result.FontId = cell.FontId;
            result.Alignment = cell.Alignment;
            result.FillId = cell.FillId;
            result.BorderId = cell.BorderId;
            result.ApplyNumberFormat = cell.ApplyNumberFormat;
            result.ApplyBorder = cell.ApplyBorder;
            result.ApplyAlignment = cell.ApplyAlignment;
            result.ApplyFill = cell.ApplyFill;
            result.ApplyFont = cell.ApplyFont;
            return result;
            
        }
    }
}

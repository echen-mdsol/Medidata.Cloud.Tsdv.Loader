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
        private uint headerStyleIndex = 0;
        private uint textStyleIndex = 0;
        public uint GetHeaderStyleIndex()
        {
            return headerStyleIndex;
        }

        public uint GetTextStyleIndex()
        {
            return textStyleIndex;
        }

        public void AttachTo(SpreadsheetDocument doc)
        {
            //If anything wrong happens, then swith to default styleindex
            try
            {
                var styleSheet = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;
                uint headerStyleId =
                    styleSheet.CellStyles.Descendants<CellStyle>()
                        .First(o => o.Name == Enum.GetName(typeof (CellStyleIdentifier), CellStyleIdentifier.Output))
                        .FormatId;

                uint textStyleId =
                    styleSheet.CellStyles.Descendants<CellStyle>()
                        .First(o => o.Name == Enum.GetName(typeof (CellStyleIdentifier), CellStyleIdentifier.Normal))
                        .FormatId;

                var list = styleSheet.CellStyleFormats.Descendants<CellFormat>().ToList();

                var headerItem = list[checked((int) headerStyleId)];
                var textStyleItem = list[checked((int) textStyleId)];

                //You cannot append a copied cell format by the restriction of Openxml sdk, you need to copy to make a new one. And you need to store the styleindex yourself.
                int newIndex = styleSheet.CellFormats.Count();
                styleSheet.CellFormats.AppendChild(CopyCellFormat(headerItem));
                headerStyleIndex = (uint) newIndex;

                styleSheet.CellFormats.AppendChild(CopyCellFormat(textStyleItem));
                newIndex++;
                textStyleIndex = (uint) newIndex;
            }
            catch
            {
                headerStyleIndex = 0;
                textStyleIndex = 0;
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

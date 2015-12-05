using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ICellStyleProvider
    {
        uint GetStyleIndex(SpreadsheetDocument doc, string styleName);
    }
}
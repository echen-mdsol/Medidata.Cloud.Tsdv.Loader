using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ICellStyleProvider
    {
        uint GetStyleIndex(string styleName);
        void AttachTo(SpreadsheetDocument doc);
    }
}
using DocumentFormat.OpenXml;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class SpreadsheetAttributeHelper
    {
        private const string Prefix = "mdsol";
        private const string MdsolNamespace = "http://www.msdol.com/rave/tsdv/1.0";
        private const string SheetNameAttributeName = "Name";

        public static OpenXmlAttribute CreateSheetNameAttribute(string sheetName)
        {
            return new OpenXmlAttribute(Prefix, SheetNameAttributeName, MdsolNamespace, sheetName);
        }
    }
}
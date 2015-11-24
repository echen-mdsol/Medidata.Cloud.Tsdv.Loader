using DocumentFormat.OpenXml;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class SpreadsheetAttributeHelper
    {
        public const string Prefix = "mdsol";
        public const string MdsolNamespace = "http://www.msdol.com/rave/tsdv/1.0";
        public const string SheetNameAttributeName = "Name";

        public static OpenXmlAttribute CreateSheetNameAttribute(string sheetName)
        {
            return new OpenXmlAttribute(Prefix, SheetNameAttributeName, MdsolNamespace, sheetName);
        }
    }
}
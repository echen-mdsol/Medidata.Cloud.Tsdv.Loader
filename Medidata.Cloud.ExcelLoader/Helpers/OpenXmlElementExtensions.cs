using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class OpenXmlElementExtensions
    {
        private const string Prefix = "mdsol";
        private const string NamespaceUri = "http://www.msdol.com/rave/tsdv/1.0";

        public static T AddMdsolAttribute<T>(this T element, string name, string value) where T : OpenXmlElement
        {
            var attribute = new OpenXmlAttribute(Prefix, name, NamespaceUri, value);
            element.SetAttribute(attribute);
            return element;
        }

        public static string GetMdsolAttribute<T>(this T element, string name) where T : OpenXmlElement
        {
            var attribute = element.GetAttribute(name, NamespaceUri);
            return attribute.Value;
        }

        public static void AddMdsolNamespaceDeclaration(this Worksheet target)
        {
            target.AddNamespaceDeclaration(Prefix, NamespaceUri);
        }
    }
}
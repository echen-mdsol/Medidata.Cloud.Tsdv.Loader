using DocumentFormat.OpenXml;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public static class OpenXmlElementExtensions
    {
        private const string Prefix = "mdsol";
        private const string MdsolNamespace = "http://www.msdol.com/rave/tsdv/1.0";

        public static T AddMdsolAttribute<T>(this T element, string name, string value) where T : OpenXmlElement
        {
            var attribute = new OpenXmlAttribute(Prefix, name, MdsolNamespace, value);
            element.SetAttribute(attribute);
            return element;
        }

        public static string GetMdsolAttribute<T>(this T element, string name) where T : OpenXmlElement
        {
            var attribute = element.GetAttribute(name, MdsolNamespace);
            return attribute.Value;
        }
    }
}
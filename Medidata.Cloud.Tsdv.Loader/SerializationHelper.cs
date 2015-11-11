using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace Medidata.Cloud.Tsdv.Loader
{
    public static class SerializationHelper
    {
        public static MemoryStream ToExcelStream<T>(this List<T> list, string xsltFilePath) where T : class, new()
        {
            //XsltArgumentList args = new XsltArgumentList();
            XmlDocument listXml = list.ToXml();
            XslCompiledTransform transform = new XslCompiledTransform();
            var xslt = File.ReadAllText(xsltFilePath);
            XmlDocument xsltDoc = new XmlDocument();
            xsltDoc.LoadXml(xslt);
            transform.Load(xsltDoc);
            MemoryStream ms = new MemoryStream();
            transform.Transform(listXml,null, ms);
            return ms;
        }


        public static string ToExcel<T>(this List<T> list, string xsltFilePath) where T : class, new()
        {
            //XsltArgumentList args = new XsltArgumentList();
            var sb = new StringBuilder();
            var sw = new StringWriter(sb);

            XmlDocument listXml = list.ToXml();
            XslCompiledTransform transform = new XslCompiledTransform();
            var xslt = File.ReadAllText(xsltFilePath);
            XmlDocument xsltDoc = new XmlDocument();
            xsltDoc.LoadXml(xslt);
            transform.Load(xsltDoc);
            MemoryStream ms = new MemoryStream();
            transform.Transform(listXml, null, sw);
            return sw.ToString();
        }


        public static XmlDocument ToXml<T>(this List<T> list ) where T : class, new()
        {
            XmlDocument xmlDoc = new XmlDocument();
            XPathNavigator nav = xmlDoc.CreateNavigator();
            XmlWriterSettings writerSettings = new XmlWriterSettings();
            writerSettings.CheckCharacters = false; // could be 0x... in data - let Excel deal with that

            try
            {
                using( XmlWriter writer = XmlWriter.Create( nav.AppendChild(), writerSettings ) )
                {
                    XmlSerializer ser = new XmlSerializer( typeof( List<T> ) ); // arrayOf.. can do -> , new XmlRootAttribute("TheRootElementName"));
                    ser.Serialize( writer, list );
                }
            }
            catch
            {
                // bad characters in data. process empty list
                List<T> emptyList = new List<T>();
                nav = xmlDoc.CreateNavigator();
                using( XmlWriter writer = XmlWriter.Create( nav.AppendChild(), writerSettings ) )
                {
                    XmlSerializer seri = new XmlSerializer( typeof( List<T> ), new XmlRootAttribute( typeof( T ).Name ) );
                    seri.Serialize( writer, emptyList );
                }
            }
            return xmlDoc;
        }
    }
    
}

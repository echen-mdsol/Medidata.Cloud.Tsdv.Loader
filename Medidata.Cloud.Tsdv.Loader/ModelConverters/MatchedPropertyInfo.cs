using System.ComponentModel;

namespace Medidata.Cloud.Tsdv.Loader.ModelConverters
{
    internal class MatchedPropertyInfo
    {
        public PropertyDescriptor InterfaceProperty { get; set; }
        public PropertyDescriptor ModelProperty { get; set; }
    }
}
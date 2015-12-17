using System.Dynamic;

namespace Medidata.Cloud.ExcelLoader.Helpers
{
    public interface IExtraProperty
    {
        ExpandoObject ExtraProperties { get; }
    }
}
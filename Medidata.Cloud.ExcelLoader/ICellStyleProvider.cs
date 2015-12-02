using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ICellStyleProvider
    {
        uint GetHeaderStyleIndex();
        uint GetTextStyleIndex();
        void AttachTo(SpreadsheetDocument doc);
    }
}

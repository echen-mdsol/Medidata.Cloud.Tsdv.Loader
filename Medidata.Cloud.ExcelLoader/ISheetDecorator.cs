using OfficeOpenXml;

namespace Medidata.Cloud.ExcelLoader
{
    public interface ISheetDecorator
    {
        void Decorate(ExcelWorksheet sheet);
    }
}
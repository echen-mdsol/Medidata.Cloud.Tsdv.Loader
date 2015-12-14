namespace Medidata.Cloud.ExcelLoader.Validations
{
    public interface IValidator
    {
        IValidationResult Validate(IExcelParser excelParser);
    }
}
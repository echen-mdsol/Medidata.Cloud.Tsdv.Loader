namespace Medidata.Cloud.ExcelLoader.Validations
{
    public interface IValidationMessage
    {
        IValidationRule ByWhom { get; }
        string Message { get; }
        string Where { get; }
    }
}
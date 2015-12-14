namespace Medidata.Cloud.ExcelLoader.Validations
{
    public interface IValidationRuleResult
    {
        IValidationMessage Message { get; }
        bool ShouldContinue { get; }
    }
}
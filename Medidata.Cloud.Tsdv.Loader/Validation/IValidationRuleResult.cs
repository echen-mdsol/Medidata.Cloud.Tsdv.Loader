namespace Medidata.Cloud.Tsdv.Loader.Validation
{
    public interface IValidationRuleResult
    {
        IValidationMessage Message { get; }
        bool ShouldContinue { get; }
    }
}
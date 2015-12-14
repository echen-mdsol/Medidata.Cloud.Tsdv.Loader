namespace Medidata.Cloud.ExcelLoader.Validations.Rules
{
    internal class ValidationRuleResult : IValidationRuleResult
    {
        public IValidationMessage Message { get; set; }
        public bool ShouldContinue { get; set; }
    }
}
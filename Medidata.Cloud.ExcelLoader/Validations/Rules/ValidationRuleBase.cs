using System;

namespace Medidata.Cloud.ExcelLoader.Validations.Rules
{
    public abstract class ValidationRuleBase : IValidationRule
    {
        public virtual IValidationRuleResult Check(IExcelParser excelParser)
        {
            var shouldContinue = true;
            Action next = () => { shouldContinue = true; };
            IValidationMessage message;
            Validate(excelParser, out message, next);
            return new ValidationRuleResult {Message = message, ShouldContinue = shouldContinue};
        }

        protected abstract void Validate(IExcelParser blockPlan, out IValidationMessage message, Action next);
    }
}
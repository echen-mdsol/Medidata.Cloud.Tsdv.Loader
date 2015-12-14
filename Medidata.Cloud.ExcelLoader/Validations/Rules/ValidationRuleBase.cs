using System;

namespace Medidata.Cloud.ExcelLoader.Validations.Rules
{
    public abstract class ValidationRuleBase : IValidationRule
    {
        public virtual IValidationRuleResult Check(IExcelParser blockPlan)
        {
            var shouldContinue = true;
            Action next = () => { shouldContinue = true; };
            IValidationMessage message;
            Validate(blockPlan, out message, next);
            return new ValidationRuleResult {Message = message, ShouldContinue = shouldContinue};
        }

        protected abstract void Validate(IExcelParser blockPlan, out IValidationMessage message, Action next);
    }
}
using Medidata.Interfaces.TSDV;

namespace Medidata.Cloud.Tsdv.Loader.Validation
{
    public interface IValidationRule
    {
        IValidationRuleResult Check(IBlockPlan blockPlan);
    }
}
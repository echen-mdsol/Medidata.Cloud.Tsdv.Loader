using Medidata.Interfaces.TSDV;

namespace Medidata.Cloud.Tsdv.Loader.Validation
{
    public interface IValidator
    {
        IValidationResult Validate(IBlockPlan blockPlan);
    }
}
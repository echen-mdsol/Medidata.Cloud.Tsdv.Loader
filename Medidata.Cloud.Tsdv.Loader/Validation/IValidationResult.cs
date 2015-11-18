using System.Collections.Generic;
using Medidata.Interfaces.TSDV;

namespace Medidata.Cloud.Tsdv.Loader.Validation
{
    public interface IValidationResult
    {
        IBlockPlan ValidationTarget { get; }
        IList<IValidationMessage> Messages { get; }
    }
}
using System.Linq;
using Medidata.Cloud.Tsdv.Loader.Validation;
using Medidata.Interfaces.TSDV;
using Ploeh.AutoFixture;
using Rhino.Mocks;

namespace Medidata.Cloud.Tsdv.Loader.Tests
{
    public class ValidatorSampleCode
    {
        public void Sample()
        {
            var fixture = new Fixture();

            // Target block plan to verify.
            var blockPlan = fixture.Create<IBlockPlan>();

            // Creates several validation rules.
            var validationRule1 = fixture.Create<IValidationRule>();
            var validationRule2 = fixture.Create<IValidationRule>();
            var validationRule3 = fixture.Create<IValidationRule>();

            // Creates a validator with the rules.
            var validator = new DefaultSequentialRuleValidator(validationRule1, validationRule2, validationRule3);

            // Uses the validator to virify the block plan.
            var result = validator.Validate(blockPlan);

            // Gets errors and warnings from the validation result.
            var errors = result.Messages.OfType<IValidationError>();
            var warnings = result.Messages.OfType<IValidationWarning>();
        }
    }
}
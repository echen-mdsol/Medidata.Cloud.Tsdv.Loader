using Medidata.Cloud.Tsdv.Loader.Validation;

namespace Medidata.Cloud.Tsdv.Loader.Extensions
{
    public static class StringValidationMessageExtensions
    {
        public static IValidationMessage ToValidationWarning(this string target)
        {
            return new ValidationWarning(target);
        }

        public static IValidationMessage ToValidationError(this string target)
        {
            return new ValidationError(target);
        }
    }
}
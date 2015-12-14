using Medidata.Cloud.ExcelLoader.Validations;

namespace Medidata.Cloud.ExcelLoader.Helpers
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
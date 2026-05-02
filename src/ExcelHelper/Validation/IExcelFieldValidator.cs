namespace ExcelHelper.Validation;

/// <summary>
///     Validates a field value during reading or writing.
/// </summary>
public interface IExcelFieldValidator
{
    /// <summary>
    ///     Validates the specified field value.
    /// </summary>
    /// <param name="args">The validation arguments.</param>
    /// <returns>A <see cref="ValidationResult" /> indicating whether validation succeeded.</returns>
    ValidationResult Validate(ValidateArgs args);
}

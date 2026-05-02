namespace ExcelHelper.Validation;

using System;

/// <summary>
///     Internal validator that wraps a boolean predicate.
/// </summary>
internal sealed class FuncValidator : IExcelFieldValidator
{
    private readonly Func<ValidateArgs, bool> _predicate;

    /// <summary>
    ///     Initializes a new instance of the <see cref="FuncValidator" /> class.
    /// </summary>
    /// <param name="predicate">The predicate that returns <c>true</c> if valid.</param>
    public FuncValidator(Func<ValidateArgs, bool> predicate)
    {
        _predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
    }

    /// <summary>
    ///     Validates using the wrapped predicate.
    /// </summary>
    public ValidationResult Validate(ValidateArgs args)
    {
        return _predicate(args)
            ? ValidationResult.Success()
            : ValidationResult.Failed("Validation failed.");
    }
}

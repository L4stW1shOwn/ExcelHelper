namespace ExcelHelper.Validation;

/// <summary>
///     Represents the result of a validation operation.
/// </summary>
public sealed class ValidationResult
{
    private ValidationResult(bool isValid, string? errorMessage)
    {
        IsValid = isValid;
        ErrorMessage = errorMessage;
    }

    /// <summary>
    ///     Gets a value indicating whether validation succeeded.
    /// </summary>
    public bool IsValid { get; }

    /// <summary>
    ///     Gets the error message if validation failed.
    /// </summary>
    /// <remarks>
    ///     This message may contain raw user data originating from the Excel file.
    ///     Ensure proper encoding (e.g., HTML encoding) before rendering in a web UI.
    /// </remarks>
    public string? ErrorMessage { get; }

    /// <summary>
    ///     Creates a successful validation result.
    /// </summary>
    /// <returns>A successful <see cref="ValidationResult" />.</returns>
    public static ValidationResult Success()
    {
        return new ValidationResult(true, null);
    }

    /// <summary>
    ///     Creates a failed validation result with the specified error message.
    /// </summary>
    /// <param name="errorMessage">The error message. May contain user-supplied data; encode before displaying in a UI.</param>
    /// <returns>A failed <see cref="ValidationResult" />.</returns>
    public static ValidationResult Failed(string errorMessage)
    {
        return new ValidationResult(false, errorMessage);
    }
}
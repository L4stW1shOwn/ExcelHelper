using System;

namespace ExcelHelper.Core;

/// <summary>
///     Event arguments for the <see cref="ExcelConfiguration.ValidationFailed" /> delegate.
/// </summary>
public sealed class ValidationFailedEventArgs : ExcelEventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ValidationFailedEventArgs" /> class.
    /// </summary>
    /// <param name="context">The context.</param>
    /// <param name="fieldName">The name of the field that failed validation, or null for record validation.</param>
    /// <param name="fieldValue">The value that failed validation.</param>
    /// <param name="errorMessage">The validation error message.</param>
    public ValidationFailedEventArgs(ExcelContext context, string? fieldName, object? fieldValue, string errorMessage)
        : base(context)
    {
        FieldName = fieldName;
        FieldValue = fieldValue;
        ErrorMessage = errorMessage ?? throw new ArgumentNullException(nameof(errorMessage));
    }

    /// <summary>
    ///     Gets the name of the field that failed validation.
    /// </summary>
    public string? FieldName { get; }

    /// <summary>
    ///     Gets the value that failed validation.
    /// </summary>
    public object? FieldValue { get; }

    /// <summary>
    ///     Gets the validation error message.
    /// </summary>
    public string ErrorMessage { get; }
}

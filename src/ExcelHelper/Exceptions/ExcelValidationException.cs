using System;
using ExcelHelper.Core;

namespace ExcelHelper.Exceptions;

/// <summary>
///     Exception thrown when a validation rule fails during reading or writing.
/// </summary>
public class ExcelValidationException : ExcelHelperException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelValidationException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="fieldName">The name of the field that failed validation.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    public ExcelValidationException(string message, string? fieldName, int row, int column)
        : base(message)
    {
        FieldName = fieldName;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelValidationException" /> class with a reference to the inner
    ///     exception.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="fieldName">The name of the field that failed validation.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="innerException">The inner exception.</param>
    public ExcelValidationException(string message, string? fieldName, int row, int column, Exception innerException)
        : base(message, innerException)
    {
        FieldName = fieldName;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelValidationException" /> class with context.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="fieldName">The name of the field that failed validation.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="context">The context associated with the exception.</param>
    public ExcelValidationException(string message, string? fieldName, int row, int column, ExcelContext? context)
        : base(message, context)
    {
        FieldName = fieldName;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelValidationException" /> class with context and inner exception.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="fieldName">The name of the field that failed validation.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="context">The context associated with the exception.</param>
    /// <param name="innerException">The inner exception.</param>
    public ExcelValidationException(string message, string? fieldName, int row, int column, ExcelContext? context, Exception innerException)
        : base(message, context, innerException)
    {
        FieldName = fieldName;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Gets the name of the field that failed validation.
    /// </summary>
    public string? FieldName { get; }

    /// <summary>
    ///     Gets the 1-based row index where the validation failed.
    /// </summary>
    public int Row { get; }

    /// <summary>
    ///     Gets the 1-based column index where the validation failed.
    /// </summary>
    public int Column { get; }
}

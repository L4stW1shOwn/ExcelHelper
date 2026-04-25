using System;

namespace ExcelHelper.Exceptions;

/// <summary>
///     Exception thrown when a type conversion fails.
/// </summary>
public class ExcelTypeConversionException : ExcelHelperException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelTypeConversionException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="fieldValue">The raw cell value.</param>
    /// <param name="targetType">The target type.</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    public ExcelTypeConversionException(
        string message,
        object? fieldValue,
        Type targetType,
        string? fieldName,
        int row,
        int column)
        : base(message)
    {
        FieldValue = fieldValue;
        TargetType = targetType;
        FieldName = fieldName;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelTypeConversionException" /> class with a reference to the inner
    ///     exception.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="fieldValue">The raw cell value.</param>
    /// <param name="targetType">The target type.</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="innerException">The inner exception.</param>
    public ExcelTypeConversionException(
        string message,
        object? fieldValue,
        Type targetType,
        string? fieldName,
        int row,
        int column,
        Exception innerException)
        : base(message, innerException)
    {
        FieldValue = fieldValue;
        TargetType = targetType;
        FieldName = fieldName;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Gets the name of the field being converted.
    /// </summary>
    public string? FieldName { get; }

    /// <summary>
    ///     Gets the raw value from the cell.
    /// </summary>
    public object? FieldValue { get; }

    /// <summary>
    ///     Gets the target type of the conversion.
    /// </summary>
    public Type TargetType { get; }

    /// <summary>
    ///     Gets the 1-based row index.
    /// </summary>
    public int Row { get; }

    /// <summary>
    ///     Gets the 1-based column index.
    /// </summary>
    public int Column { get; }
}
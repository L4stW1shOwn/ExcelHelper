using ExcelHelper.Core;

namespace ExcelHelper.Validation;

/// <summary>
///     Represents the arguments passed to a field validator.
/// </summary>
public sealed class ValidateArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ValidateArgs" /> class.
    /// </summary>
    /// <param name="value">The field value.</param>
    /// <param name="fieldName">The name of the field.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="context">The current Excel context.</param>
    public ValidateArgs(object? value, string? fieldName, int row, int column, ExcelContext? context)
    {
        Value = value;
        FieldName = fieldName;
        Row = row;
        Column = column;
        Context = context;
    }

    /// <summary>
    ///     Gets the field value.
    /// </summary>
    public object? Value { get; }

    /// <summary>
    ///     Gets the name of the field.
    /// </summary>
    public string? FieldName { get; }

    /// <summary>
    ///     Gets the 1-based row index.
    /// </summary>
    public int Row { get; }

    /// <summary>
    ///     Gets the 1-based column index.
    /// </summary>
    public int Column { get; }

    /// <summary>
    ///     Gets the current Excel context.
    /// </summary>
    public ExcelContext? Context { get; }
}

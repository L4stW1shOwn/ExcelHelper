namespace ExcelHelper.Core;

/// <summary>
///     Event arguments for the <see cref="ExcelConfiguration.BadDataFound" /> delegate.
/// </summary>
public sealed class BadDataFoundEventArgs : ExcelEventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="BadDataFoundEventArgs" /> class.
    /// </summary>
    /// <param name="context">The reading context.</param>
    /// <param name="fieldName">The name of the field where bad data was found.</param>
    /// <param name="rawValue">The raw cell value.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="column">The 1-based column index.</param>
    public BadDataFoundEventArgs(ReadingContext context, string? fieldName, object? rawValue, int row, int column)
        : base(context)
    {
        FieldName = fieldName;
        RawValue = rawValue;
        Row = row;
        Column = column;
    }

    /// <summary>
    ///     Gets the reading context.
    /// </summary>
    public ReadingContext ReadingContext => (ReadingContext)Context;

    /// <summary>
    ///     Gets the name of the field where bad data was found.
    /// </summary>
    public string? FieldName { get; }

    /// <summary>
    ///     Gets the raw cell value.
    /// </summary>
    public object? RawValue { get; }

    /// <summary>
    ///     Gets the 1-based row index.
    /// </summary>
    public int Row { get; }

    /// <summary>
    ///     Gets the 1-based column index.
    /// </summary>
    public int Column { get; }
}

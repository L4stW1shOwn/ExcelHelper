namespace ExcelHelper.Core;

/// <summary>
///     Event arguments for the <see cref="ExcelConfiguration.MissingFieldFound" /> delegate.
/// </summary>
public sealed class MissingFieldEventArgs : ExcelEventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="MissingFieldEventArgs" /> class.
    /// </summary>
    /// <param name="context">The reading context.</param>
    /// <param name="fieldName">The name of the missing field.</param>
    /// <param name="row">The 1-based row index.</param>
    public MissingFieldEventArgs(ReadingContext context, string? fieldName, int row)
        : base(context)
    {
        FieldName = fieldName;
        Row = row;
    }

    /// <summary>
    ///     Gets the reading context.
    /// </summary>
    public ReadingContext ReadingContext => (ReadingContext)Context;

    /// <summary>
    ///     Gets the name of the missing field.
    /// </summary>
    public string? FieldName { get; }

    /// <summary>
    ///     Gets the 1-based row index where the field was missing.
    /// </summary>
    public int Row { get; }
}

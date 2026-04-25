namespace ExcelHelper.Core;

/// <summary>
///     Provides context information during writing operations.
/// </summary>
public sealed class WritingContext : ExcelContext
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="WritingContext" /> class.
    /// </summary>
    /// <param name="configuration">The configuration for the writing operation.</param>
    public WritingContext(ExcelConfiguration configuration)
        : base(configuration)
    {
    }

    /// <summary>
    ///     Gets a value indicating whether the header has already been written.
    /// </summary>
    public bool HasHeaderBeenWritten { get; internal set; }
}
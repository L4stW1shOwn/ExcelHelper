using System;
using ExcelHelper.Exceptions;

namespace ExcelHelper.Core;

/// <summary>
///     Event arguments for the <see cref="ExcelConfiguration.ReadingExceptionOccurred" /> delegate.
/// </summary>
public sealed class ReadingExceptionEventArgs : ExcelEventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ReadingExceptionEventArgs" /> class.
    /// </summary>
    /// <param name="context">The reading context.</param>
    /// <param name="exception">The exception that occurred.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> or <paramref name="exception" /> is null.</exception>
    public ReadingExceptionEventArgs(ReadingContext context, ExcelHelperException exception)
        : base(context)
    {
        Exception = exception ?? throw new ArgumentNullException(nameof(exception));
    }

    /// <summary>
    ///     Gets the reading context.
    /// </summary>
    public ReadingContext ReadingContext => (ReadingContext)Context;

    /// <summary>
    ///     Gets the exception that occurred during reading.
    /// </summary>
    public ExcelHelperException Exception { get; }
}

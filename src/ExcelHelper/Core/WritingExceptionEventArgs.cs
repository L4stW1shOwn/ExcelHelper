using System;
using ExcelHelper.Exceptions;

namespace ExcelHelper.Core;

/// <summary>
///     Event arguments for the <see cref="ExcelConfiguration.WritingExceptionOccurred" /> delegate.
/// </summary>
public sealed class WritingExceptionEventArgs : ExcelEventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="WritingExceptionEventArgs" /> class.
    /// </summary>
    /// <param name="context">The writing context.</param>
    /// <param name="exception">The exception that occurred.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> or <paramref name="exception" /> is null.</exception>
    public WritingExceptionEventArgs(WritingContext context, ExcelHelperException exception)
        : base(context)
    {
        Exception = exception ?? throw new ArgumentNullException(nameof(exception));
    }

    /// <summary>
    ///     Gets the writing context.
    /// </summary>
    public WritingContext WritingContext => (WritingContext)Context;

    /// <summary>
    ///     Gets the exception that occurred during writing.
    /// </summary>
    public ExcelHelperException Exception { get; }
}

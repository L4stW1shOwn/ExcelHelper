using System;

namespace ExcelHelper.Core;

/// <summary>
///     Base class for all event arguments used by ExcelHelper delegates.
/// </summary>
public abstract class ExcelEventArgs : EventArgs
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelEventArgs" /> class.
    /// </summary>
    /// <param name="context">The context associated with the event.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> is null.</exception>
    protected ExcelEventArgs(ExcelContext context)
    {
        Context = context ?? throw new ArgumentNullException(nameof(context));
    }

    /// <summary>
    ///     Gets the context associated with the event.
    /// </summary>
    public ExcelContext Context { get; }
}

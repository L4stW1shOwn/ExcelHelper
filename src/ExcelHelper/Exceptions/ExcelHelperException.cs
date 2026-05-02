using System;
using ExcelHelper.Core;

namespace ExcelHelper.Exceptions;

/// <summary>
///     Base exception for all ExcelHelper errors.
/// </summary>
public class ExcelHelperException : Exception
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelHelperException" /> class.
    /// </summary>
    public ExcelHelperException()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelHelperException" /> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public ExcelHelperException(string message)
        : base(message)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelHelperException" /> class with a specified error message and a
    ///     reference to the inner exception that is the cause of this exception.
    /// </summary>
    /// <param name="message">The error message that explains the reason for the exception.</param>
    /// <param name="innerException">
    ///     The exception that is the cause of the current exception, or a null reference if no inner
    ///     exception is specified.
    /// </param>
    public ExcelHelperException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelHelperException" /> class with a specified error message and
    ///     context.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="context">The context associated with the exception.</param>
    public ExcelHelperException(string message, ExcelContext? context)
        : base(message)
    {
        Context = context;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelHelperException" /> class with a specified error message, context,
    ///     and inner exception.
    /// </summary>
    /// <param name="message">The error message that explains the reason for the exception.</param>
    /// <param name="context">The context associated with the exception.</param>
    /// <param name="innerException">The inner exception.</param>
    public ExcelHelperException(string message, ExcelContext? context, Exception innerException)
        : base(message, innerException)
    {
        Context = context;
    }

    /// <summary>
    ///     Gets the context associated with the exception, if available.
    ///     Cast to <see cref="ReadingContext" /> or <see cref="WritingContext" /> as needed.
    /// </summary>
    /// <remarks>
    ///     This is a live reference to the mutable context at the time the exception was thrown.
    ///     Do not cache it for later inspection, as its state may change during subsequent operations.
    /// </remarks>
    public ExcelContext? Context { get; }
}

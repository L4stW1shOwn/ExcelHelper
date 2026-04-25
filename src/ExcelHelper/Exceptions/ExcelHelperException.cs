using System;

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
}
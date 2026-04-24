using System;

namespace ExcelHelper.Exceptions
{
    /// <summary>
    /// Exception thrown when a mapping configuration is invalid.
    /// </summary>
    public class ExcelMappingException : ExcelHelperException
    {
        /// <summary>
        /// Gets the type being mapped when the error occurred.
        /// </summary>
        public Type? MappedType { get; }

        /// <summary>
        /// Gets the member name being mapped when the error occurred.
        /// </summary>
        public string? MemberName { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMappingException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="mappedType">The type being mapped.</param>
        /// <param name="memberName">The member name being mapped.</param>
        public ExcelMappingException(string message, Type? mappedType, string? memberName)
            : base(message)
        {
            MappedType = mappedType;
            MemberName = memberName;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMappingException"/> class with a reference to the inner exception.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="mappedType">The type being mapped.</param>
        /// <param name="memberName">The member name being mapped.</param>
        /// <param name="innerException">The inner exception.</param>
        public ExcelMappingException(string message, Type? mappedType, string? memberName, Exception innerException)
            : base(message, innerException)
        {
            MappedType = mappedType;
            MemberName = memberName;
        }
    }
}

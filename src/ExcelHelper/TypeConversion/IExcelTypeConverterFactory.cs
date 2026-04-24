using System;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Creates instances of <see cref="IExcelTypeConverter{T}"/>.
    /// </summary>
    public interface IExcelTypeConverterFactory
    {
        /// <summary>
        /// Creates a type converter for the specified type.
        /// </summary>
        /// <param name="type">The type to convert.</param>
        /// <returns>The type converter instance.</returns>
        object CreateConverter(Type type);
    }
}

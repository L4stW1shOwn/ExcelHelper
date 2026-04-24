using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts an Excel cell value to and from a .NET type.
    /// </summary>
    /// <typeparam name="T">The .NET type to convert.</typeparam>
    public interface IExcelTypeConverter<T>
    {
        /// <summary>
        /// Converts a cell value to the target .NET type.
        /// </summary>
        /// <param name="value">The cell value.</param>
        /// <param name="cultureInfo">The culture to use for conversion.</param>
        /// <returns>The converted value.</returns>
        T ConvertFromExcel(object? value, CultureInfo cultureInfo);

        /// <summary>
        /// Converts a .NET value to an Excel-compatible value.
        /// </summary>
        /// <param name="value">The .NET value.</param>
        /// <param name="cultureInfo">The culture to use for conversion.</param>
        /// <returns>The Excel-compatible value.</returns>
        object? ConvertToExcel(object? value, CultureInfo cultureInfo);
    }
}

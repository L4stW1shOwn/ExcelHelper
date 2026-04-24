using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts between Excel cell values and <see cref="string"/>.
    /// </summary>
    public sealed class StringConverter : IExcelTypeConverter<string>
    {
        /// <summary>
        /// Converts a cell value to a string.
        /// </summary>
        public string ConvertFromExcel(object? value, CultureInfo cultureInfo)
        {
            if (value == null)
                return string.Empty;

            return value.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Converts a string to an Excel-compatible value.
        /// </summary>
        public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
        {
            return value;
        }
    }
}

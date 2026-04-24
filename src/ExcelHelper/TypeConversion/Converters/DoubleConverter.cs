using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts between Excel cell values and <see cref="double"/>.
    /// </summary>
    public sealed class DoubleConverter : IExcelTypeConverter<double>
    {
        /// <summary>
        /// Converts a cell value to a double.
        /// </summary>
        public double ConvertFromExcel(object? value, CultureInfo cultureInfo)
        {
            if (value == null)
                return default;

            if (value is double d)
                return d;

            if (value is int i)
                return i;

            if (value is long l)
                return l;

            if (value is decimal dec)
                return (double)dec;

            if (value is float f)
                return f;

            if (value is string s && double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, cultureInfo, out var result))
                return result;

            return Convert.ToDouble(value, cultureInfo);
        }

        /// <summary>
        /// Converts a double to an Excel-compatible value.
        /// </summary>
        public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
        {
            return value;
        }
    }
}

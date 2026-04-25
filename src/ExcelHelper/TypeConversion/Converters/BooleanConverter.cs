using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Converts between Excel cell values and <see cref="bool" />.
/// </summary>
public sealed class BooleanConverter : IExcelTypeConverter<bool>
{
    /// <summary>
    ///     Converts a cell value to a boolean.
    /// </summary>
    public bool ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return default;
        }

        if (value is bool b)
        {
            return b;
        }

        if (value is string s)
        {
            if (bool.TryParse(s, out var result))
            {
                return result;
            }

            var trimmed = s.Trim();
            if (string.Equals(trimmed, "1", StringComparison.Ordinal) ||
                string.Equals(trimmed, "yes", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "y", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(trimmed, "0", StringComparison.Ordinal) ||
                string.Equals(trimmed, "no", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "false", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "n", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (value is int i)
        {
            return i != 0;
        }

        if (value is double d)
        {
            return d != 0;
        }

        return Convert.ToBoolean(value, cultureInfo);
    }

    /// <summary>
    ///     Converts a boolean to an Excel-compatible value.
    /// </summary>
    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        return value;
    }
}
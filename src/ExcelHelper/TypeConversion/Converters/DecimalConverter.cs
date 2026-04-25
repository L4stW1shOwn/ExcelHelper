using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Converts between Excel cell values and <see cref="decimal" />.
/// </summary>
public sealed class DecimalConverter : IExcelTypeConverter<decimal>
{
    /// <summary>
    ///     Converts a cell value to a decimal.
    /// </summary>
    public decimal ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return default;
        }

        if (value is decimal dec)
        {
            return dec;
        }

        if (value is double d)
        {
            return (decimal)d;
        }

        if (value is int i)
        {
            return i;
        }

        if (value is long l)
        {
            return l;
        }

        if (value is string s && decimal.TryParse(s, NumberStyles.Number, cultureInfo, out var result))
        {
            return result;
        }

        return Convert.ToDecimal(value, cultureInfo);
    }

    /// <summary>
    ///     Converts a decimal to an Excel-compatible value.
    /// </summary>
    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        return value;
    }
}
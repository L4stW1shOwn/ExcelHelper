using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Converts between Excel cell values and <see cref="long" />.
/// </summary>
public sealed class Int64Converter : IExcelTypeConverter<long>
{
    /// <summary>
    ///     Converts a cell value to a long.
    /// </summary>
    public long ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return default;
        }

        if (value is long l)
        {
            return l;
        }

        if (value is int i)
        {
            return i;
        }

        if (value is double d)
        {
            return (long)d;
        }

        if (value is decimal dec)
        {
            return (long)dec;
        }

        if (value is string s && long.TryParse(s, NumberStyles.Integer, cultureInfo, out var result))
        {
            return result;
        }

        return Convert.ToInt64(value, cultureInfo);
    }

    /// <summary>
    ///     Converts a long to an Excel-compatible value.
    /// </summary>
    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        return value;
    }
}
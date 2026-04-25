using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Converts between Excel cell values and <see cref="int" />.
/// </summary>
public sealed class Int32Converter : IExcelTypeConverter<int>
{
    /// <summary>
    ///     Converts a cell value to an integer.
    /// </summary>
    public int ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return default;
        }

        if (value is int i)
        {
            return i;
        }

        if (value is double d)
        {
            return (int)d;
        }

        if (value is decimal dec)
        {
            return (int)dec;
        }

        if (value is long l)
        {
            return (int)l;
        }

        if (value is string s && int.TryParse(s, NumberStyles.Integer, cultureInfo, out var result))
        {
            return result;
        }

        return Convert.ToInt32(value, cultureInfo);
    }

    /// <summary>
    ///     Converts an integer to an Excel-compatible value.
    /// </summary>
    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        return value;
    }
}
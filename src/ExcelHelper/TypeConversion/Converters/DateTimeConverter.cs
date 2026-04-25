using System;
using System.Globalization;
using ExcelHelper.Internal;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Converts between Excel cell values and <see cref="DateTime" />.
/// </summary>
public sealed class DateTimeConverter : IExcelTypeConverter<DateTime>
{
    /// <summary>
    ///     Converts a cell value to a <see cref="DateTime" />.
    /// </summary>
    public DateTime ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return default;
        }

        if (value is DateTime dt)
        {
            return dt;
        }

        if (value is double d)
        {
            return OADateConverter.FromOADate(d);
        }

        if (value is string s)
        {
            if (DateTime.TryParse(s, cultureInfo, DateTimeStyles.None, out var result))
            {
                return result;
            }

            if (double.TryParse(s, NumberStyles.Float, cultureInfo, out var oaDate))
            {
                return OADateConverter.FromOADate(oaDate);
            }
        }

        return Convert.ToDateTime(value, cultureInfo);
    }

    /// <summary>
    ///     Converts a <see cref="DateTime" /> to an Excel-compatible value.
    /// </summary>
    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        if (value is DateTime dt)
        {
            return OADateConverter.ToOADate(dt);
        }

        return value;
    }
}
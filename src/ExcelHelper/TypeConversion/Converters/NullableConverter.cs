using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Converts between Excel cell values and <typeparamref name="T" />.
/// </summary>
/// <typeparam name="T">The underlying non-nullable type.</typeparam>
public sealed class NullableConverter<T> : IExcelTypeConverter<T?>
    where T : struct
{
    private readonly IExcelTypeConverter<T> _innerConverter;

    /// <summary>
    ///     Initializes a new instance of the <see cref="NullableConverter{T}" /> class.
    /// </summary>
    /// <param name="innerConverter">The converter for the underlying type.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="innerConverter" /> is null.</exception>
    public NullableConverter(IExcelTypeConverter<T> innerConverter)
    {
        _innerConverter = innerConverter ?? throw new ArgumentNullException(nameof(innerConverter));
    }

    /// <summary>
    ///     Converts a cell value to a nullable value.
    /// </summary>
    public T? ConvertFromExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return null;
        }

        if (value is string s && string.IsNullOrWhiteSpace(s))
        {
            return null;
        }

        return _innerConverter.ConvertFromExcel(value, cultureInfo);
    }

    /// <summary>
    ///     Converts a nullable value to an Excel-compatible value.
    /// </summary>
    public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
    {
        if (value == null)
        {
            return null;
        }

        return _innerConverter.ConvertToExcel(value, cultureInfo);
    }
}
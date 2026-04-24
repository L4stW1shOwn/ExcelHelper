using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts between Excel cell values and <typeparamref name="TEnum"/>.
    /// </summary>
    /// <typeparam name="TEnum">The enum type.</typeparam>
    public sealed class EnumConverter<TEnum> : IExcelTypeConverter<TEnum>
        where TEnum : struct, Enum
    {
        /// <summary>
        /// Converts a cell value to an enum.
        /// </summary>
        public TEnum ConvertFromExcel(object? value, CultureInfo cultureInfo)
        {
            if (value == null)
                return default;

            if (value is TEnum e)
                return e;

            if (value is string s)
            {
                if (Enum.TryParse<TEnum>(s, true, out var result))
                    return result;

                // Try parsing numeric string
                if (int.TryParse(s, NumberStyles.Integer, cultureInfo, out var intValue))
                {
                    if (Enum.IsDefined(typeof(TEnum), intValue))
                        return (TEnum)Enum.ToObject(typeof(TEnum), intValue);
                }
            }

            if (value is int i && Enum.IsDefined(typeof(TEnum), i))
                return (TEnum)Enum.ToObject(typeof(TEnum), i);

            throw new InvalidOperationException($"Cannot convert '{value}' to enum '{typeof(TEnum).Name}'.");
        }

        /// <summary>
        /// Converts an enum to an Excel-compatible value.
        /// </summary>
        public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
        {
            if (value is TEnum e)
                return e.ToString();

            return value;
        }
    }
}

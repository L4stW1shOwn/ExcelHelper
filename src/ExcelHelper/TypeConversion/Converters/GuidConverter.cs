using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts between Excel cell values and <see cref="Guid"/>.
    /// </summary>
    public sealed class GuidConverter : IExcelTypeConverter<Guid>
    {
        /// <summary>
        /// Converts a cell value to a <see cref="Guid"/>.
        /// </summary>
        public Guid ConvertFromExcel(object? value, CultureInfo cultureInfo)
        {
            if (value == null)
                return default;

            if (value is Guid g)
                return g;

            if (value is string s && Guid.TryParse(s, out var result))
                return result;

            throw new InvalidOperationException($"Cannot convert '{value}' to Guid.");
        }

        /// <summary>
        /// Converts a <see cref="Guid"/> to an Excel-compatible value.
        /// </summary>
        public object? ConvertToExcel(object? value, CultureInfo cultureInfo)
        {
            if (value is Guid g)
                return g.ToString();

            return value;
        }
    }
}

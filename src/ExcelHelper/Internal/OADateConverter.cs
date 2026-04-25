using System;

namespace ExcelHelper.Internal;

/// <summary>
///     Converts between Excel OADate serial numbers and <see cref="DateTime" />.
/// </summary>
public static class OADateConverter
{
    /// <summary>
    ///     Converts an OADate serial number to a <see cref="DateTime" />.
    /// </summary>
    /// <param name="oaDate">The OADate serial number.</param>
    /// <returns>The <see cref="DateTime" />.</returns>
    public static DateTime FromOADate(double oaDate)
    {
        return DateTime.FromOADate(oaDate);
    }

    /// <summary>
    ///     Converts a <see cref="DateTime" /> to an OADate serial number.
    /// </summary>
    /// <param name="dateTime">The <see cref="DateTime" />.</param>
    /// <returns>The OADate serial number.</returns>
    public static double ToOADate(DateTime dateTime)
    {
        return dateTime.ToOADate();
    }
}
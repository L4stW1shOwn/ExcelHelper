using System;
using System.Globalization;
using ExcelHelper.Exceptions;
using ExcelHelper.Mapping;

namespace ExcelHelper.Core;

/// <summary>
///     Configuration settings for ExcelHelper read and write operations.
/// </summary>
public class ExcelConfiguration
{
    /// <summary>
    ///     Gets the collection of registered class maps.
    /// </summary>
    public ExcelClassMapCollection Maps { get; } = new();

    /// <summary>
    ///     Registers a class map for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to map.</typeparam>
    /// <param name="map">The class map.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="map" /> is null.</exception>
    public void RegisterClassMap<T>(ExcelClassMap<T> map)
    {
        if (map == null)
        {
            throw new ArgumentNullException(nameof(map));
        }

        Maps.Add(map);
    }

    /// <summary>
    ///     Registers a class map by creating an instance of the specified map type.
    /// </summary>
    /// <typeparam name="TClassMap">The type of the class map to create.</typeparam>
    public void RegisterClassMap<TClassMap>()
        where TClassMap : ExcelClassMap, new()
    {
        var map = new TClassMap();
        Maps.Add(map);
    }

    /// <summary>
    ///     Gets or sets the object resolver used to create instances during mapping.
    /// </summary>
    public IObjectResolver ObjectResolver { get; set; } = new DefaultObjectResolver();

    /// <summary>
    ///     Gets or sets the culture used for type conversions.
    ///     Default is <see cref="System.Globalization.CultureInfo.InvariantCulture" />.
    /// </summary>
    public CultureInfo CultureInfo { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    ///     Gets or sets a value indicating whether the Excel file has a header record.
    ///     Default is true.
    /// </summary>
    public bool HasHeaderRecord { get; set; } = true;

    /// <summary>
    ///     Gets or sets the 1-based row index of the header.
    ///     Default is 1.
    /// </summary>
    public int HeaderRow { get; set; } = 1;

    /// <summary>
    ///     Gets or sets the 1-based row index where data starts.
    ///     Default is 2.
    /// </summary>
    public int StartRow { get; set; } = 2;

    /// <summary>
    ///     Gets or sets the name of the worksheet to read from or write to.
    ///     If null, the first worksheet (index 0) is used.
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    ///     Gets or sets the 0-based index of the worksheet to use when <see cref="SheetName" /> is not specified.
    ///     Default is 0.
    /// </summary>
    public int SheetIndex { get; set; } = 0;

    /// <summary>
    ///     Gets or sets a value indicating whether to interpret Excel OADate values when converting to DateTime.
    ///     Default is true.
    /// </summary>
    public bool UseOADate { get; set; } = true;

    /// <summary>
    ///     Gets or sets a value indicating whether to trim whitespace from cell values.
    ///     Default is false.
    /// </summary>
    public bool TrimCellValues { get; set; } = false;

    /// <summary>
    ///     Gets or sets a value indicating whether blank rows should be ignored during reading.
    ///     Default is true.
    /// </summary>
    public bool IgnoreBlankRows { get; set; } = true;

    /// <summary>
    ///     Gets or sets a value indicating whether to ignore Excel error references such as #REF!, #VALUE!, etc.
    ///     Default is false.
    /// </summary>
    public bool IgnoreReferences { get; set; } = false;

    /// <summary>
    ///     Gets or sets a function to prepare header names for matching (e.g., trimming, lower-casing).
    ///     Default trims whitespace.
    /// </summary>
    public Func<string, string> PrepareHeaderForMatch { get; set; } = static header => header?.Trim() ?? string.Empty;

    /// <summary>
    ///     Gets or sets a callback that is invoked when an exception occurs during reading.
    ///     Return true to ignore the exception and continue; false to throw.
    /// </summary>
    public Func<ReadingExceptionEventArgs, bool> ReadingExceptionOccurred { get; set; } = static _ => false;

    /// <summary>
    ///     Gets or sets a callback that is invoked when an exception occurs during writing.
    ///     Return true to ignore the exception and continue; false to throw.
    /// </summary>
    public Func<WritingExceptionEventArgs, bool> WritingExceptionOccurred { get; set; } = static _ => false;

    /// <summary>
    ///     Gets or sets a callback that is invoked when a field is missing during reading.
    /// </summary>
    public Action<MissingFieldEventArgs> MissingFieldFound { get; set; } = static _ => { };

    /// <summary>
    ///     Gets or sets a callback that is invoked when bad data is found in a cell.
    ///     Return true to ignore and continue; false to throw.
    /// </summary>
    public Func<BadDataFoundEventArgs, bool> BadDataFound { get; set; } = static _ => false;

    /// <summary>
    ///     Gets or sets a callback invoked when any field or record validation fails
    ///     (i.e. when an <see cref="Validation.IExcelFieldValidator" /> or <see cref="Validation.IExcelRecordValidator{T}" /> returns invalid).
    ///     Return true to ignore the validation error and continue; false to throw.
    /// </summary>
    public Func<ValidationFailedEventArgs, bool> ValidationFailed { get; set; } = static _ => false;
}

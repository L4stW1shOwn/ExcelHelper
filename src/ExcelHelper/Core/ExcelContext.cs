using System;
using ExcelHelper.Mapping;
using OfficeOpenXml;

namespace ExcelHelper.Core;

/// <summary>
///     Provides context information during Excel read and write operations.
/// </summary>
public abstract class ExcelContext
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelContext" /> class.
    /// </summary>
    /// <param name="configuration">The configuration for the operation.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="configuration" /> is null.</exception>
    protected ExcelContext(ExcelConfiguration configuration)
    {
        Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
    }

    /// <summary>
    ///     Gets the configuration used for the current operation.
    /// </summary>
    public ExcelConfiguration Configuration { get; }

    /// <summary>
    ///     Gets the worksheet currently being processed.
    /// </summary>
    public ExcelWorksheet? Worksheet { get; internal set; }

    /// <summary>
    ///     Gets the 1-based current row index.
    /// </summary>
    public int Row { get; internal set; }

    /// <summary>
    ///     Gets the 1-based current column index.
    /// </summary>
    public int Column { get; internal set; }

    /// <summary>
    ///     Gets the name of the current worksheet, or null if not set.
    /// </summary>
    public string? SheetName => Worksheet?.Name;

    /// <summary>
    ///     Registers a class map for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to map.</typeparam>
    /// <param name="map">The class map.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="map" /> is null.</exception>
    internal void RegisterClassMap<T>(ExcelClassMap<T> map)
    {
        if (map == null)
        {
            throw new ArgumentNullException(nameof(map));
        }

        Configuration.Maps.Add(map);
    }

    /// <summary>
    ///     Registers a class map by creating an instance of the specified map type.
    /// </summary>
    /// <typeparam name="TClassMap">The type of the class map to create.</typeparam>
    internal void RegisterClassMap<TClassMap>()
        where TClassMap : ExcelClassMap, new()
    {
        var map = new TClassMap();
        Configuration.Maps.Add(map);
    }

    /// <summary>
    ///     Removes all registered class maps.
    /// </summary>
    internal void UnregisterClassMap()
    {
        Configuration.Maps.Clear();
    }

    /// <summary>
    ///     Removes the class map for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to remove the map for.</typeparam>
    internal void UnregisterClassMap<T>()
    {
        Configuration.Maps.Remove<T>();
    }
}
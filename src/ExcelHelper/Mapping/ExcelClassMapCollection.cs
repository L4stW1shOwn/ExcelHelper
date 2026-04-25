using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelHelper.Mapping;

/// <summary>
///     A collection of <see cref="ExcelClassMap{T}" /> instances.
/// </summary>
public class ExcelClassMapCollection : IEnumerable<KeyValuePair<Type, object>>
{
    private readonly Dictionary<Type, object> _maps = new();

    /// <summary>
    ///     Gets an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator.</returns>
    public IEnumerator<KeyValuePair<Type, object>> GetEnumerator()
    {
        return _maps.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    ///     Registers a class map for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to map.</typeparam>
    /// <param name="map">The class map.</param>
    public void Add<T>(ExcelClassMap<T> map)
    {
        _maps[typeof(T)] = map ?? throw new ArgumentNullException(nameof(map));
    }

    /// <summary>
    ///     Gets the class map for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to get the map for.</typeparam>
    /// <returns>The class map, or null if not found.</returns>
    public ExcelClassMap<T>? Get<T>()
    {
        if (_maps.TryGetValue(typeof(T), out var map))
        {
            return (ExcelClassMap<T>)map;
        }

        return null;
    }

    /// <summary>
    ///     Determines whether a class map exists for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to check.</typeparam>
    /// <returns><c>true</c> if a map exists; otherwise, <c>false</c>.</returns>
    public bool Contains<T>()
    {
        return _maps.ContainsKey(typeof(T));
    }

    /// <summary>
    ///     Removes the class map for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to remove the map for.</typeparam>
    /// <returns><c>true</c> if the map was removed; otherwise, <c>false</c>.</returns>
    public bool Remove<T>()
    {
        return _maps.Remove(typeof(T));
    }

    /// <summary>
    ///     Registers a class map.
    /// </summary>
    /// <param name="map">The class map.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="map" /> is null.</exception>
    public void Add(ExcelClassMap map)
    {
        if (map == null)
        {
            throw new ArgumentNullException(nameof(map));
        }

        _maps[map.ClassType] = map;
    }

    /// <summary>
    ///     Removes all class maps from the collection.
    /// </summary>
    public void Clear()
    {
        _maps.Clear();
    }
}
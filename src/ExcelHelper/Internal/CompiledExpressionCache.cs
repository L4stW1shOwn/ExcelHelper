using System;
using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelHelper.Internal;

/// <summary>
///     Caches compiled expressions for property getters and setters.
/// </summary>
public static class CompiledExpressionCache
{
    private static readonly ConcurrentDictionary<(Type, string), Delegate> Getters = new();
    private static readonly ConcurrentDictionary<(Type, string), Delegate> Setters = new();

    /// <summary>
    ///     Gets a compiled getter for the specified property.
    /// </summary>
    /// <typeparam name="T">The type containing the property.</typeparam>
    /// <typeparam name="TMember">The property type.</typeparam>
    /// <param name="propertyName">The name of the property.</param>
    /// <returns>A compiled getter delegate.</returns>
    public static Func<T, TMember> GetGetter<T, TMember>(string propertyName)
    {
        if (propertyName == null)
        {
            throw new ArgumentNullException(nameof(propertyName));
        }

        var key = (typeof(T), propertyName);
        var del = Getters.GetOrAdd(key, k =>
        {
            var property = typeof(T).GetProperty(k.Item2, BindingFlags.Public | BindingFlags.Instance);
            if (property == null)
            {
                throw new InvalidOperationException($"Property '{k.Item2}' not found on type '{k.Item1.Name}'.");
            }

            var instance = Expression.Parameter(k.Item1, "instance");
            var propertyAccess = Expression.Property(instance, property);
            var lambda = Expression.Lambda<Func<T, TMember>>(propertyAccess, instance);
            return lambda.Compile();
        });

        return (Func<T, TMember>)del;
    }

    /// <summary>
    ///     Gets a compiled setter for the specified property.
    /// </summary>
    /// <typeparam name="T">The type containing the property.</typeparam>
    /// <typeparam name="TMember">The property type.</typeparam>
    /// <param name="propertyName">The name of the property.</param>
    /// <returns>A compiled setter delegate.</returns>
    public static Action<T, TMember> GetSetter<T, TMember>(string propertyName)
    {
        if (propertyName == null)
        {
            throw new ArgumentNullException(nameof(propertyName));
        }

        var key = (typeof(T), propertyName);
        var del = Setters.GetOrAdd(key, k =>
        {
            var property = typeof(T).GetProperty(k.Item2, BindingFlags.Public | BindingFlags.Instance);
            if (property == null)
            {
                throw new InvalidOperationException($"Property '{k.Item2}' not found on type '{k.Item1.Name}'.");
            }

            var instance = Expression.Parameter(k.Item1, "instance");
            var value = Expression.Parameter(typeof(TMember), "value");
            var propertyAccess = Expression.Property(instance, property);

            // Handle type conversion for nullable types
            var assign = Expression.Assign(propertyAccess, value);
            var lambda = Expression.Lambda<Action<T, TMember>>(assign, instance, value);
            return lambda.Compile();
        });

        return (Action<T, TMember>)del;
    }

    /// <summary>
    ///     Clears all cached expressions.
    /// </summary>
    public static void Clear()
    {
        Getters.Clear();
        Setters.Clear();
    }
}
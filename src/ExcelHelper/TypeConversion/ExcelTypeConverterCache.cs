using System;
using System.Collections.Concurrent;
using ExcelHelper.Internal;

namespace ExcelHelper.TypeConversion;

/// <summary>
///     Caches instances of <see cref="IExcelTypeConverter{T}" />.
/// </summary>
public sealed class ExcelTypeConverterCache
{
    private static readonly ConcurrentDictionary<Type, object> Converters = new();

    /// <summary>
    ///     Gets the converter for the specified type.
    /// </summary>
    /// <typeparam name="T">The type to convert.</typeparam>
    /// <returns>The converter instance.</returns>
    public IExcelTypeConverter<T> GetConverter<T>()
    {
        var converter = Converters.GetOrAdd(typeof(T), _ => CreateConverter(typeof(T)));
        return (IExcelTypeConverter<T>)converter;
    }

    /// <summary>
    ///     Gets the converter for the specified type.
    /// </summary>
    /// <param name="type">The type to convert.</param>
    /// <returns>The converter instance.</returns>
    public object GetConverter(Type type)
    {
        return type == null
            ? throw new ArgumentNullException(nameof(type))
            : Converters.GetOrAdd(type, _ => CreateConverter(type));
    }

    private static object CreateConverter(Type type)
    {
        var underlying = ReflectionHelper.GetUnderlyingType(type);

        if (type == typeof(string))
        {
            return new StringConverter();
        }

        if (type == typeof(int) || type == typeof(int?))
        {
            return new Int32Converter();
        }

        if (type == typeof(long) || type == typeof(long?))
        {
            return new Int64Converter();
        }

        if (type == typeof(double) || type == typeof(double?))
        {
            return new DoubleConverter();
        }

        if (type == typeof(decimal) || type == typeof(decimal?))
        {
            return new DecimalConverter();
        }

        if (type == typeof(bool) || type == typeof(bool?))
        {
            return new BooleanConverter();
        }

        if (type == typeof(DateTime) || type == typeof(DateTime?))
        {
            return new DateTimeConverter();
        }

        if (type == typeof(Guid) || type == typeof(Guid?))
        {
            return new GuidConverter();
        }

        if (underlying.IsEnum)
        {
            var converterType = typeof(EnumConverter<>).MakeGenericType(underlying);
            return Activator.CreateInstance(converterType)!;
        }

        if (ReflectionHelper.IsNullable(type))
        {
            var genericType = Nullable.GetUnderlyingType(type);
            if (genericType != null)
            {
                var converterType = typeof(NullableConverter<>).MakeGenericType(genericType);
                var innerConverter = CreateConverter(genericType);
                return Activator.CreateInstance(converterType, innerConverter)!;
            }
        }

        throw new NotSupportedException($"No converter registered for type '{type.FullName}'.");
    }

    /// <summary>
    ///     Clears all cached converters.
    /// </summary>
    public static void Clear()
    {
        Converters.Clear();
    }
}
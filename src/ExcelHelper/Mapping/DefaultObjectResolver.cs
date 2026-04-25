using System;

namespace ExcelHelper.Mapping;

/// <summary>
///     Default implementation of <see cref="IObjectResolver" /> that uses Activator.CreateInstance.
/// </summary>
public sealed class DefaultObjectResolver : IObjectResolver
{
    /// <summary>
    ///     Creates an instance of the specified type using <see cref="Activator.CreateInstance(Type)" />.
    /// </summary>
    /// <param name="type">The type to create.</param>
    /// <returns>A new instance of the specified type.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="type" /> is null.</exception>
    public object Resolve(Type type)
    {
        return type == null ? throw new ArgumentNullException(nameof(type)) : Activator.CreateInstance(type)!;
    }
}
using System;

namespace ExcelHelper.Mapping
{
    /// <summary>
    /// Resolves instances of types during mapping operations.
    /// </summary>
    public interface IObjectResolver
    {
        /// <summary>
        /// Creates an instance of the specified type.
        /// </summary>
        /// <param name="type">The type to create.</param>
        /// <returns>A new instance of the specified type.</returns>
        object Resolve(Type type);
    }
}

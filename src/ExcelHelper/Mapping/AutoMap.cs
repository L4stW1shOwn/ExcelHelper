using System;
using System.Linq;
using System.Reflection;
using ExcelHelper.Internal;

namespace ExcelHelper.Mapping
{
    /// <summary>
    /// Automatically maps class properties based on reflection and attributes.
    /// </summary>
    public static class AutoMapper
    {
        /// <summary>
        /// Automatically maps all public properties of <typeparamref name="T"/> to the specified <see cref="ExcelClassMap{T}"/>.
        /// </summary>
        /// <typeparam name="T">The type to auto-map.</typeparam>
        /// <param name="map">The class map to populate.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="map"/> is null.</exception>
        public static void Apply<T>(ExcelClassMap<T> map)
        {
            if (map == null)
                throw new ArgumentNullException(nameof(map));

            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.CanWrite);

            int currentIndex = 0;

            foreach (var property in properties)
            {
                var ignoreAttr = property.GetCustomAttribute<ExcelIgnoreAttribute>();
                if (ignoreAttr != null)
                    continue;

                var columnAttr = property.GetCustomAttribute<ExcelColumnAttribute>();
                var indexAttr = property.GetCustomAttribute<ExcelIndexAttribute>();
                var nameAttr = property.GetCustomAttribute<ExcelNameAttribute>();
                var defaultAttr = property.GetCustomAttribute<ExcelDefaultAttribute>();

                var memberMapData = new MemberMapData
                {
                    Member = property,
                    Name = nameAttr?.Name ?? columnAttr?.Name ?? property.Name,
                    Index = indexAttr?.Index ?? columnAttr?.Index ?? -1,
                    Default = defaultAttr?.Value ?? columnAttr?.Default,
                    TypeConverter = columnAttr?.Converter,
                    Ignore = false,
                    IsOptional = false
                };

                // If no index is specified, auto-increment
                if (memberMapData.Index < 0)
                {
                    memberMapData.Index = currentIndex;
                    currentIndex++;
                }
                else
                {
                    currentIndex = memberMapData.Index + 1;
                }

                map.MemberMaps.Add(memberMapData);
            }
        }
    }
}

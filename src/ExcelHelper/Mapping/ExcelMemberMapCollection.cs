using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelHelper.Mapping
{
    /// <summary>
    /// A collection of <see cref="ExcelMemberMap{TClass, TMember}"/> instances.
    /// </summary>
    public class ExcelMemberMapCollection : IEnumerable<MemberMapData>
    {
        private readonly List<MemberMapData> _maps = new List<MemberMapData>();

        /// <summary>
        /// Gets the number of member maps in the collection.
        /// </summary>
        public int Count => _maps.Count;

        /// <summary>
        /// Adds a member map data to the collection.
        /// </summary>
        /// <param name="map">The member map data to add.</param>
        public void Add(MemberMapData map)
        {
            if (map == null)
                throw new ArgumentNullException(nameof(map));

            _maps.Add(map);
        }

        /// <summary>
        /// Finds a member map by column name.
        /// </summary>
        /// <param name="name">The column name.</param>
        /// <returns>The member map data, or null if not found.</returns>
        public MemberMapData? FindByName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;

            return _maps.FirstOrDefault(m =>
                !m.Ignore &&
                !string.IsNullOrEmpty(m.Name) &&
                string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Finds a member map by column index.
        /// </summary>
        /// <param name="index">The column index.</param>
        /// <returns>The member map data, or null if not found.</returns>
        public MemberMapData? FindByIndex(int index)
        {
            return _maps.FirstOrDefault(m => !m.Ignore && m.Index == index);
        }

        /// <summary>
        /// Gets an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator.</returns>
        public IEnumerator<MemberMapData> GetEnumerator()
        {
            return _maps.Where(m => !m.Ignore).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}

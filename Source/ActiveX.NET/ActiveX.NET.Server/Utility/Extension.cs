using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Server.Utility
{
    public static class Extension
    {
        public static void TryForEach<T>(this IEnumerable<T> enumerable, Action<T, int> handler)
        {
            if (enumerable == null || handler == null)
            {
                return;
            }

            for (int index = 0; index < enumerable.Count(); ++index)
            {
                handler(enumerable.ElementAt(index), index);
            }
        }

        public static void TryForEach<T>(this IEnumerable<T> enumerable, Action<T> handler)
        {
            if (enumerable == null || handler == null)
            {
                return;
            }

            for (int index = 0; index < enumerable.Count(); ++index)
            {
                handler(enumerable.ElementAt(index));
            }
        }

        // <summary>
        /// Determines whether the collection is null or contains no elements.
        /// </summary>
        /// <typeparam name="T">The IEnumerable type.</typeparam>
        /// <param name="enumerable">The enumerable, which may be null or empty.</param>
        /// <returns>
        ///     <c>true</c> if the IEnumerable is null or empty; otherwise, <c>false</c>.
        /// </returns>
        public static bool IsNullOrEmpty<T>(this IEnumerable<T> enumerable)
        {
            if (enumerable == null)
            {
                return true;
            }
            /* If this is a list, use the Count property for efficiency. 
             * The Count property is O(1) while IEnumerable.Count() is O(N). */
            var collection = enumerable as ICollection<T>;
            if (collection != null)
            {
                return collection.Count < 1;
            }
            return !enumerable.Any();
        }
    }
};

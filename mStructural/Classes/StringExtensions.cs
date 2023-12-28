using System;

namespace mStructural.Classes
{
    public static class StringExtensions
    {
        // extension method for string to take in extra parameter which is string comparison enum
        public static bool Contains(this string source, string tocheck, StringComparison comp)
        {
            return source?.IndexOf(tocheck, comp) >= 0;
        }
    }
}

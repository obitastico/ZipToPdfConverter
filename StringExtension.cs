using System;

namespace ZipToPdfConverter
{
    public static class StringExtension
    {
        public static string RemoveRightToChar(this string str, string chr)
        {
            int charIndex = str.IndexOf(chr, StringComparison.Ordinal);
            if (charIndex >= 0)
                return str.Substring(0, charIndex);

            return str;
        }
    }
}
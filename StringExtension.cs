using System;
using System.IO;

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

        public static string UpdateFileName(this string str)
        {
            var semester =
                $"[Q1{(DateTime.Compare(new DateTime(2023, 8, 27), DateTime.Now) > 0 ? 1 : 2)} - {(DateTime.Now.Month > 8 || DateTime.Now.Month < 2 ? 1 : 2)}]";
            return str.Replace(Path.GetFileName(str), $"{DateTime.Now:yy-MM-dd} l {Path.GetFileNameWithoutExtension(str)} {semester}.pdf");
        }
    }
}
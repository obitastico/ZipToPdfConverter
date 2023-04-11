using System;
using System.IO;

namespace ZipToPdfConverter
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            string docPath = args.Length > 0 ? args[0] : @"D:\Downloads\4 Weimars Krisenjahre-20230322.zip";
            FileConverter fileConverter = new FileConverter();
            fileConverter.ConvertZipToPdf(docPath);
        }
    }
}
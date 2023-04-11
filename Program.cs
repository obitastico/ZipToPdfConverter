namespace ZipToPdfConverter
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            string docPath = args.Length > 0 ? args[0] : @"D:\Downloads\1 Der Versailler Vertrag-20230301.zip";
            FileConverter fileConverter = new FileConverter();
            fileConverter.ConvertZipToPdf(docPath);
        }
    }
}
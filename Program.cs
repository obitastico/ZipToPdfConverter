namespace ZipToPdfConverter
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                FileConverter fileConverter = new FileConverter();
                fileConverter.ConvertZipToPdf(args[0]);
            }
            
        }
    }
}
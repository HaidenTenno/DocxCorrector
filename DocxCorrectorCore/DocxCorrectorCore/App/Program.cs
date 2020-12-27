namespace DocxCorrectorCore.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            System.Console.OutputEncoding = System.Text.Encoding.UTF8;
            CommandLineParser.Parse(args);
        }
    }
}

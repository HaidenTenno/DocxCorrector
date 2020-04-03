using System;

namespace DocxCorrectorCore.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider();

            //featuresProvider.GenerateHeadersFootersInfoJSON(Models.HeaderFooterType.Footer, Config.DocFilePath, Config.HeadersFootersInfoFilePath);

            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }
    }
}

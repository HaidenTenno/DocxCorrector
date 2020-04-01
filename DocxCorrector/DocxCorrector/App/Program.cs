using System;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider(type: FeaturesProviderType.Spire);
            
            featuresProvider.GenerateSectionsPropertiesJSON(Config.DocFilePath, Config.PagesPropertiesFilePath);
            
            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }
    }
}

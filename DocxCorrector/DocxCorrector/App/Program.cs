using System;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = FeaturesProvider.GetInstance(type: FeaturesProviderType.InteropMultipleApp);

            //featuresProvider.GenerateCSVFiles(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileName);
            //featuresProvider.GeneratePagesPropertiesJSON(Config.DocFilePath, Config.PagesPropertiesFilePath);

            featuresProvider.TestCorrectorSpeed(Config.FilesToInpectDirectoryPath);

            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }
    }
}

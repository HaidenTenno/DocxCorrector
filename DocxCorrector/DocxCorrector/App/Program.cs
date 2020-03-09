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

            Console.WriteLine("OLD");
            TimeCounter.CountTime(() => featuresProvider.GenerateNormalizedCSVFiles(Config.FilesToInpectDirectoryPath, Config.NormalizedPropertiesFileName));
            Console.WriteLine("\nNEW");
            TimeCounter.CountTime(() => featuresProvider.GenerateNormalizedCSVFilesAsync(Config.FilesToInpectDirectoryPath, Config.NormalizedPropertiesFileNameAstnc));

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}

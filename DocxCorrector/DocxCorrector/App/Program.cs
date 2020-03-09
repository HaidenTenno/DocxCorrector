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
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsync(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileName));
            Console.WriteLine("\nNEW");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsyncWithAsyncFilesIteration(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileNameAsync));

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}

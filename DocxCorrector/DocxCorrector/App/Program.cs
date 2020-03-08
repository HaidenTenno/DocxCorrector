using System;
using System.Threading;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = FeaturesProvider.GetInstance(type: FeaturesProviderType.Interop);

            Console.WriteLine("OLD");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsync(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileName));
            Console.WriteLine("\nNEW");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsync1(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileNameAsync));

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}

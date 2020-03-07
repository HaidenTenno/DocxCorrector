using System;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = FeaturesProvider.GetInstance(type: FeaturesProviderType.Interop);

            featuresProvider.GeneratePagesPropertiesJSON(filePath: Config.DocFilePath, resultFilePath: Config.PagesPropertiesFilePath);

            //TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsync(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileNameAsync));

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}

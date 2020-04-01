﻿using System;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = new FeaturesProvider(type: FeaturesProviderType.GemBox);
            
            featuresProvider.GenerateSectionsPropertiesJSON(Config.DocFilePath, Config.PagesPropertiesFilePath);
            featuresProvider.GenerateCSVFiles(Config.FilesToInpectDirectoryPath, Config.ParagraphPropertiesFileName);

            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }
    }
}

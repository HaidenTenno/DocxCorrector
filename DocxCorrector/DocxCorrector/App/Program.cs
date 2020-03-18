﻿using System;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = FeaturesProvider.GetInstance(type: FeaturesProviderType.GemBox);

            featuresProvider.GenerateNormalizedCSVFiles(Config.FilesToInpectDirectoryPath, Config.NormalizedPropertiesFileName);

            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }
    }
}

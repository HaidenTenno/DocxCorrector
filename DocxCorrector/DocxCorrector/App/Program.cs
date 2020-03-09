using System;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            FeaturesProvider featuresProvider = FeaturesProvider.GetInstance(type: FeaturesProviderType.GemBox);

            Console.WriteLine("Синхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFiles(Config.FilesToInpectDirectoryPath, Config.SyncParagraphsSyncIteration));
            Console.WriteLine("\nАсинхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsync(Config.FilesToInpectDirectoryPath, Config.AsyncParagraphsSyncIteration));
            Console.WriteLine("\nCинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesWithAsyncFilesIteration(Config.FilesToInpectDirectoryPath, Config.SyncParagraphsAsyncIteration));
            Console.WriteLine("\nАсинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.CountTime(() => featuresProvider.GenerateCSVFilesAsyncWithAsyncFilesIteration(Config.FilesToInpectDirectoryPath, Config.AsyncParagraphsAsyncIteration));

            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }
    }
}

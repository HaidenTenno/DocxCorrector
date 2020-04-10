using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using DocxCorrectorCore.Services.Corrector;
using DocxCorrectorCore.Services.PropertiesPuller;
using DocxCorrectorCore.Services;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.App
{
    // Функции программы, доступные глобально
    public sealed class FeaturesProvider
    {
        // Private
        private readonly Corrector Corrector;
        private readonly PropertiesPuller PropertiesPuller;

        // Public
        public FeaturesProvider()
        {
            Corrector = new CorrectorGemBox();
            PropertiesPuller = new PropertiesPullerGemBox();
        }

        // Напечатать содержимое всех параграфов документа filePath
        public void PrintParagraphs(string filePath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            Corrector.PrintAllParagraphs(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Проанализировать документ filePath и Создать JSON файл в директории resultDirPath со свойствами его страниц
        public void GeneratePagesPropertiesJSON(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<PageProperties> pagesProperties = PropertiesPuller.GetAllPagesProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string pagesPropertiesJSON = JSONMaker.MakeJSON(pagesProperties);
            string resultFilePath = Path.Combine(resultDirPath, Config.PagesPropertiesFileName);
            FileWriter.WriteToFile(resultFilePath, pagesPropertiesJSON);
        }

        // Проанализировать документ filePath и Создать JSON файл в директории resultDirPath со свойствами его секций
        public void GenerateSectionsPropertiesJSON(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<SectionProperties> sectionsProperties = PropertiesPuller.GetAllSectionsProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string sectionsPropertiesJSON = JSONMaker.MakeJSON(sectionsProperties);
            string resultFilePath = Path.Combine(resultDirPath, Config.SectionsPropertiesFileName);
            FileWriter.WriteToFile(resultFilePath, sectionsPropertiesJSON);
        }

        // Проанализировать документ filePath и создать JSON файл в директории resultDirPath со свойствами колонтитулов типа type
        public void GenerateHeadersFootersInfoJSON(HeaderFooterType type, string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<HeaderFooterInfo> headersFootersInfo = PropertiesPuller.GetHeadersFootersInfo(type: type, filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string headersFootersInfoJSON = JSONMaker.MakeJSON(headersFootersInfo);
            string resultFilePath = Path.Combine(resultDirPath, Config.HeadersFootersInfoFileName);
            FileWriter.WriteToFile(resultFilePath, headersFootersInfoJSON);
        }

        // Пройтись по всем поддиректориям rootDir и в каждой создать csv файл, где будут записаны свойства параграфов для всех docx файлов в этой директории
        public void GenerateCSVFiles(string rootDir)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = PropertiesPuller.GetAllParagraphsProperties(filePath: filePath);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                });

                string resultFilePath = Path.Combine(subDir, Config.ParagraphsPropertiesFileName);
                FileWriter.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // Пройтись по всем поддиректориям rootDir и в каждой создать csv файл, где будут записаны нормализованные свойства параграфов для всех docx файлов в этой директории
        public void GenerateNormalizedCSVFiles(string rootDir)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<NormalizedProperties> normalizedPropertiesForDir = new List<NormalizedProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<NormalizedProperties> normalizedPropertiesForFile = PropertiesPuller.GetNormalizedProperties(filePath: filePath);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    normalizedPropertiesForDir.AddRange(normalizedPropertiesForFile);
                });

                string resultFilePath = Path.Combine(subDir, Config.NormalizedPropertiesFileName);
                FileWriter.FillCSV(resultFilePath, normalizedPropertiesForDir);
            });
        }

        // GenerateCSVFiles, основанный на асинхронном методе
        public void GenerateCSVFilesAsync(string rootDir)
        {
            IPropertiesPullerAsync? asyncPuller = PropertiesPuller as IPropertiesPullerAsync;

            if (asyncPuller == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = asyncPuller.GetAllParagraphsPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                });

                string resultFilePath = Path.Combine(subDir, Config.AsyncParagraphsSyncIterationFileName);
                FileWriter.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // GenerateCSVFiles с асинхронным анализом файлов
        public void GenerateCSVFilesWithAsyncFilesIteration(string rootDir)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                Task.WaitAll(DirectoryIterator.IterateDocxFilesAsync(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = PropertiesPuller.GetAllParagraphsProperties(filePath: filePath);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                }));

                string resultFilePath = Path.Combine(subDir, Config.SyncParagraphsAsyncIterationFileName);
                FileWriter.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // GenerateCSVFiles, основанный на асинхронном методе с асинхронным анализом файлов
        public void GenerateCSVFilesAsyncWithAsyncFilesIteration(string rootDir)
        {
            IPropertiesPullerAsync? asyncPuller = PropertiesPuller as IPropertiesPullerAsync;

            if (asyncPuller == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                Task.WaitAll(DirectoryIterator.IterateDocxFilesAsync(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = asyncPuller.GetAllParagraphsPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                }));

                string resultFilePath = Path.Combine(subDir, Config.AsyncParagraphsAsyncIterationFileName);
                FileWriter.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // GenerateNormalizedCSVFiles, основанный на асинхнонном методе
        public void GenerateNormalizedCSVFilesAsync(string rootDir)
        {
            IPropertiesPullerAsync? asyncPuller = PropertiesPuller as IPropertiesPullerAsync;

            if (asyncPuller == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<NormalizedProperties> normalizedPropertiesForDir = new List<NormalizedProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<NormalizedProperties> normalizedPropertiesForFile = asyncPuller.GetNormalizedPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    normalizedPropertiesForDir.AddRange(normalizedPropertiesForFile);
                });

                string resultFilePath = Path.Combine(subDir, Config.NormalizedPropertiesFileName);
                FileWriter.FillCSV(resultFilePath, normalizedPropertiesForDir);
            });
        }

        // Тест скорости работы синхронных и асинхронных методов при вытигивании свойств из документов
        public void TestCorrectorSpeed(string rootDir)
        {
            Console.WriteLine("Синхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFiles(rootDir));
            Console.WriteLine("\nАсинхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFilesAsync(rootDir));
            Console.WriteLine("\nCинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFilesWithAsyncFilesIteration(rootDir));
            Console.WriteLine("\nАсинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFilesAsyncWithAsyncFilesIteration(rootDir));
        }

        // Сохранить документ filePath как pdf в директории resultDirPath
        public void SaveDocumentAsPdf(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            FileWriter.SaveDocumentAsPdf(filePath, resultDirPath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Сохранить страницы документ filePath как отдельные pdf в директории resultDirPath
        public void SavePagesAsPdf(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            FileWriter.SavePagesAsPdf(filePath, resultDirPath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }
    }
}

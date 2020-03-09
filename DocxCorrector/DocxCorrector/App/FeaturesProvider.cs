#nullable enable
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using DocxCorrector.Services.Corrector;
using DocxCorrector.Services;
using DocxCorrector.Models;

namespace DocxCorrector.App
{
    public enum FeaturesProviderType
    {
        Interop,
        GemBox,
        InteropMultipleApp
    }

    // Функции программы, доступные глобально
    public sealed class FeaturesProvider
    {
        // Private
        private static FeaturesProvider? Instance;

        private readonly Corrector Corrector;
        
        private FeaturesProvider(Corrector corrector)
        {
            Corrector = corrector;
        }

        // Public
        // Получение экземпляра класса, реализующих возможности приложения через библиотеку type
        public static FeaturesProvider GetInstance(FeaturesProviderType type)
        {
            if (Instance == null)
            {
                Instance = type switch
                {
                    FeaturesProviderType.Interop => new FeaturesProvider(corrector: new CorrectorInterop()),
                    FeaturesProviderType.GemBox => new FeaturesProvider(corrector: new CorrectorGemBox()),
                    FeaturesProviderType.InteropMultipleApp => new FeaturesProvider(corrector: new CorrectorInteropMultipleApps()),
                    _ => throw new NotImplementedException()
                };
            }

            return Instance;
        }

        public void PrintParagraphs(string filePath)
        {
            Corrector.PrintAllParagraphs(filePath: filePath);
        }

        // Проанализировать документ filePath и Создать JSON файл resultFilePath со свойствами его страниц
        public void GeneratePagesPropertiesJSON(string filePath, string resultFilePath)
        {
            List<PageProperties> pagesProperties = Corrector.GetAllPagesProperties(filePath: filePath);
            string pagesPropertiesJSON = JSONMaker.MakeJSON(pagesProperties);
            FileWriter.WriteToFile(resultFilePath, pagesPropertiesJSON);
        }

        // Пройтись по всем поддиректориям rootDir и в каждой создать csv файл с именем resultFileName, где будут результаты для всех docx файлов в этой директории
        public void GenerateCSVFiles(string rootDir, string resultFileName)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = Corrector.GetAllParagraphsProperties(filePath: filePath);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                });

                FileWriter.FillCSV(String.Concat(subDir, resultFileName), propertiesForDir);
            });
        }

        // Получение данных для программы Ромы
        public void GenerateNormalizedCSVFiles(string rootDir, string resultFileName)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<NormalizedProperties> normalizedPropertiesForDir = new List<NormalizedProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<NormalizedProperties> normalizedPropertiesForFile = Corrector.GetNormalizedProperties(filePath: filePath);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    normalizedPropertiesForDir.AddRange(normalizedPropertiesForFile);
                });

                FileWriter.FillCSV(String.Concat(subDir, resultFileName), normalizedPropertiesForDir);
            });
        }

        // Проанализировать документ filePath и Создать JSON файл resultFilePath со списком ошибок, с учетом того, что все параграфы в документе определенного типа
        public void CheckParagraphs(string filePath, string resultFilePath)
        {
            Console.WriteLine("Введите тип проверяемых параграфов:\n0 - абзац\n1 - элемент списка\n2 - подпись к рисунку");
            string userAnswer = Console.ReadLine();
            int userAnserInt;
            bool result = int.TryParse(userAnswer, out userAnserInt);

            if (!result)
            {
                Console.WriteLine("Недопустимый ответ");
                return;
            }

            List<ParagraphResult> paragraphResults;

            switch ((ElementType)userAnserInt)
            {
                case ElementType.Paragraph:
                case ElementType.List:
                case ElementType.ImageSign:
                    paragraphResults = Corrector.GetMistakesForElementType(filePath: filePath, elementType: (ElementType)userAnserInt);
                    break;

                default:
                    Console.WriteLine("Ответ не поддерживается");
                    return;
            }

            string resultJSON = JSONMaker.MakeJSON(results: paragraphResults);
            FileWriter.WriteToFile(resultFilePath, resultJSON);
        }

        // GenerateCSVFiles, основанный на асинхронном методе
        public void GenerateCSVFilesAsync(string rootDir, string resultFileName)
        {
            ICorrecorAsync? asyncCorretor = Corrector as ICorrecorAsync;

            if (asyncCorretor == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = asyncCorretor.GetAllParagraphsPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                });

                FileWriter.FillCSV(String.Concat(subDir, resultFileName), propertiesForDir);
            });
        }

        // GenerateCSVFiles с асинхронным анализом файлов
        public void GenerateCSVFilesWithAsyncFilesIteration(string rootDir, string resultFileName)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                Task.WaitAll(DirectoryIterator.IterateDocxFilesAsync(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = Corrector.GetAllParagraphsProperties(filePath: filePath);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                }));

                FileWriter.FillCSV(String.Concat(subDir, resultFileName), propertiesForDir);
            });
        }

        // GenerateCSVFiles, основанный на асинхронном методе с асинхронным анализом файлов
        public void GenerateCSVFilesAsyncWithAsyncFilesIteration(string rootDir, string resultFileName)
        {
            ICorrecorAsync? asyncCorretor = Corrector as ICorrecorAsync;

            if (asyncCorretor == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                Task.WaitAll(DirectoryIterator.IterateDocxFilesAsync(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphProperties> propertiesForFile = asyncCorretor.GetAllParagraphsPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                }));

                FileWriter.FillCSV(String.Concat(subDir, resultFileName), propertiesForDir);
            });
        }

        // GenerateNormalizedCSVFiles, основанный на асинхнонном методе
        public void GenerateNormalizedCSVFilesAsync(string rootDir, string resultFileName)
        {
            ICorrecorAsync? asyncCorretor = Corrector as ICorrecorAsync;

            if (asyncCorretor == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<NormalizedProperties> normalizedPropertiesForDir = new List<NormalizedProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<NormalizedProperties> normalizedPropertiesForFile = asyncCorretor.GetNormalizedPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    normalizedPropertiesForDir.AddRange(normalizedPropertiesForFile);
                });

                FileWriter.FillCSV(String.Concat(subDir, resultFileName), normalizedPropertiesForDir);
            });
        }
    }
}

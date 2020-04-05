#nullable enable
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using DocxCorrectorCore.Services.Corrector;
using DocxCorrectorCore.Services;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.App
{
    // Функции программы, доступные глобально
    public sealed class FeaturesProvider
    {
        // Private
        private readonly Corrector Corrector;

        // Public
        public FeaturesProvider()
        {
            Corrector = new CorrectorGemBox();
        }

        // Напечатать содержимое всех параграфов документа filePath
        public void PrintParagraphs(string filePath)
        {
            Corrector.PrintAllParagraphs(filePath: filePath);
        }

        // Проанализировать документ filePath и Создать JSON файл resultFilePath со свойствами его страниц
        public void GeneratePagesPropertiesJSON(string filePath, string resultFilePath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<PageProperties> pagesProperties = Corrector.GetAllPagesProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string pagesPropertiesJSON = JSONMaker.MakeJSON(pagesProperties);
            FileWriter.WriteToFile(resultFilePath, pagesPropertiesJSON);
        }

        // Проанализировать документ filePath и Создать JSON файл resultFilePath со свойствами его секций
        public void GenerateSectionsPropertiesJSON(string filePath, string resultFilePath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<SectionProperties> sectionsProperties = Corrector.GetAllSectionsProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string sectionsPropertiesJSON = JSONMaker.MakeJSON(sectionsProperties);
            FileWriter.WriteToFile(resultFilePath, sectionsPropertiesJSON);
        }

        // Проанализировать документ filePath и создать JSON файл resultFilePath со свойствами колонтитулов типа type
        public void GenerateHeadersFootersInfoJSON(HeaderFooterType type, string filePath, string resultFilePath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<HeaderFooterInfo> headersFootersInfo = Corrector.GetHeadersFootersInfo(type: type, filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string headersFootersInfoJSON = JSONMaker.MakeJSON(headersFootersInfo);
            FileWriter.WriteToFile(resultFilePath, headersFootersInfoJSON);
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

                FileWriter.FillCSV(Path.Combine(subDir, resultFileName), propertiesForDir);
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

                FileWriter.FillCSV(Path.Combine(subDir, resultFileName), normalizedPropertiesForDir);
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

                FileWriter.FillCSV(Path.Combine(subDir, resultFileName), propertiesForDir);
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

                FileWriter.FillCSV(Path.Combine(subDir, resultFileName), propertiesForDir);
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

                FileWriter.FillCSV(Path.Combine(subDir, resultFileName), propertiesForDir);
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

                FileWriter.FillCSV(Path.Combine(subDir, resultFileName), normalizedPropertiesForDir);
            });
        }

        // Тест скорости работы синхронных и асинхронных методов при вытигивании свойств из документов
        public void TestCorrectorSpeed(string rootDir)
        {
            Console.WriteLine("Синхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFiles(rootDir, Config.SyncParagraphsSyncIteration));
            Console.WriteLine("\nАсинхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFilesAsync(rootDir, Config.AsyncParagraphsSyncIteration));
            Console.WriteLine("\nCинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFilesWithAsyncFilesIteration(rootDir, Config.SyncParagraphsAsyncIteration));
            Console.WriteLine("\nАсинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.CountTime(() => GenerateCSVFilesAsyncWithAsyncFilesIteration(rootDir, Config.AsyncParagraphsAsyncIteration));
        }
    }
}

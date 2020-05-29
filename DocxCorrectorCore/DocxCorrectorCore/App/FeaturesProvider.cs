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

        // Напечатать информацию о структуре документа filePath
        public void PrintStructureInfo(string filePath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            PropertiesPuller.PrintDocumentStructureInfo(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Напечатать информацию о содержании документа filePath
        public void PrintTableOfContentsInfo(string filePath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            PropertiesPuller.PrintTableOfContenstsInfo(filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Проанализировать документ filePath и Создать JSON файл в директории resultDirPath со свойствами его страниц
        public void GeneratePagesPropertiesJSON(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<PageProperties> pagesProperties = PropertiesPuller.GetAllPagesProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string pagesPropertiesJSON = JSONWorker.MakeJSON(pagesProperties);
            string resultFilePath = Path.Combine(resultDirPath, Config.PagesPropertiesFileName);
            FileWorker.WriteToFile(resultFilePath, pagesPropertiesJSON);
        }

        // Проанализировать документ filePath и Создать JSON файл в директории resultDirPath со свойствами его секций
        public void GenerateSectionsPropertiesJSON(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<SectionProperties> sectionsProperties = PropertiesPuller.GetAllSectionsProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string sectionsPropertiesJSON = JSONWorker.MakeJSON(sectionsProperties);
            string resultFilePath = Path.Combine(resultDirPath, Config.SectionsPropertiesFileName);
            FileWorker.WriteToFile(resultFilePath, sectionsPropertiesJSON);
        }

        // Проанализировать документ filePath и создать JSON файл в директории resultDirPath со свойствами колонтитулов типа type
        public void GenerateHeadersFootersInfoJSON(HeaderFooterType type, string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<HeaderFooterInfo> headersFootersInfo = PropertiesPuller.GetHeadersFootersInfo(type: type, filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string headersFootersInfoJSON = JSONWorker.MakeJSON(headersFootersInfo);
            string resultFilePath = Path.Combine(resultDirPath, Config.HeadersFootersInfoFileName);
            FileWorker.WriteToFile(resultFilePath, headersFootersInfoJSON);
        }

        // Проанализировать документ filePath и создать csv файл в директории resultDirPath со свойствами его параграфов
        public void GenerateParagraphsPropertiesCSV(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<ParagraphProperties> propertiesForFile = new List<ParagraphProperties>();
            string time = TimeCounter.GetExecutionTime(() => { propertiesForFile = PropertiesPuller.GetAllParagraphsProperties(filePath: filePath); }, TimeCounter.ResultType.TotalMilliseconds);
            Console.WriteLine($"Done {Path.GetFileName(filePath)} in {time}");
            string resultFilePath = Path.Combine(resultDirPath, Config.ParagraphsPropertiesFileName);
            FileWorker.FillCSV(resultFilePath, propertiesForFile);
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
                FileWorker.FillCSV(resultFilePath, propertiesForDir);
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
                FileWorker.FillCSV(resultFilePath, propertiesForDir);
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
                FileWorker.FillCSV(resultFilePath, propertiesForDir);
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
                FileWorker.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // Тест скорости работы синхронных и асинхронных методов при вытигивании свойств параграфов из документов
        public void TestParagraphPropertiesPullingSpeed(string rootDir)
        {
            Console.WriteLine("Синхронный анализ параграфов, синхронный проход по директории");
            TimeCounter.LogExecutionTime(() => GenerateCSVFiles(rootDir));

            // TODO: NOT SUPPORTED IN OUR DLL

            //Console.WriteLine("\nАсинхронный анализ параграфов, синхронный проход по директории");
            //TimeCounter.LogExecutionTime(() => GenerateCSVFilesAsync(rootDir));
            Console.WriteLine("\nCинхронный анализ параграфов, асинхронный проход по директории");
            TimeCounter.LogExecutionTime(() => GenerateCSVFilesWithAsyncFilesIteration(rootDir));
            //Console.WriteLine("\nАсинхронный анализ параграфов, асинхронный проход по директории");
            //TimeCounter.LogExecutionTime(() => GenerateCSVFilesAsyncWithAsyncFilesIteration(rootDir));
        }

        // Сохранить документ filePath как pdf в директории resultDirPath
        public void SaveDocumentAsPdf(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            FileWorker.SaveDocumentAsPdf(filePath, resultDirPath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Сохранить страницы документ filePath как отдельные pdf в директории resultDirPath
        public void SavePagesAsPdf(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            FileWorker.SavePagesAsPdf(filePath, resultDirPath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Вывести содержимое pdf документа filePath с помощью библиотеки GemBox.Document
        public void PrintPdfGemBoxDocument(string filePath)
        {
            if (Path.GetExtension(filePath) != ".pdf") { Console.WriteLine("Неверное расширение"); return; }
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            PropertiesPuller.PrintContent(filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Вывести содержимое pdf документа filePath с помощью библиотеки GemBox.Pdf
        public void PrintPdfGemBoxPdf(string filePath)
        {
            if (Path.GetExtension(filePath) != ".pdf") { Console.WriteLine("Неверное расширение"); return; }
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            Services.Helpers.GemBoxPdfHelper gemBoxPdfHelper = new Services.Helpers.GemBoxPdfHelper();
            gemBoxPdfHelper.PrintContent(filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
        }

        // Получить список ошибок форматирования для ВСЕГО документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        // сохранение результата в директории resultDirPath
        public void GenerateMistakesJSON(string fileToCorrect, RulesModel rules, string paragraphsClassesFile, string resultDir)
        {
            List<ClassificationResult>? paragraphsClassesList = JSONWorker.DeserializeObjectFromFile<List<ClassificationResult>>(paragraphsClassesFile);
            if (paragraphsClassesList == null) { return; }

            Console.WriteLine($"Started {Path.GetFileName(fileToCorrect)}");
            DocumentCorrections documentCorrections = new DocumentCorrections(rules);
            string time = TimeCounter.GetExecutionTime(() => { documentCorrections = Corrector.GetCorrections(fileToCorrect, rules, paragraphsClassesList); }, TimeCounter.ResultType.TotalMilliseconds);
            Console.WriteLine($"Done {Path.GetFileName(fileToCorrect)} in {time}");

            string documentCorrectionsJSON = JSONWorker.MakeJSON(documentCorrections);
            string resultFilePath = Path.Combine(resultDir, Config.MistakesFileName);
            FileWorker.WriteToFile(resultFilePath, documentCorrectionsJSON);
        }
    }
}

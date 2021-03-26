using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.BusinessLogicLayer.Corrector;
using DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller;
using DocxCorrectorCore.Services.Utilities;
using System.Linq;

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
            List<PagePropertiesGemBox> pagesProperties = PropertiesPuller.GetAllPagesProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string pagesPropertiesJSON = JSONWorker.MakeJSON(pagesProperties);
            string resultFilePath = Path.Combine(resultDirPath, DefaultFileNames.PagesPropertiesFileName);
            FileWorker.WriteToFile(resultFilePath, pagesPropertiesJSON);
        }

        // Проанализировать документ filePath и Создать JSON файл в директории resultDirPath со свойствами его секций
        public void GenerateSectionsPropertiesJSON(string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<SectionPropertiesGemBox> sectionsProperties = PropertiesPuller.GetAllSectionsProperties(filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string sectionsPropertiesJSON = JSONWorker.MakeJSON(sectionsProperties);
            string resultFilePath = Path.Combine(resultDirPath, DefaultFileNames.SectionsPropertiesFileName);
            FileWorker.WriteToFile(resultFilePath, sectionsPropertiesJSON);
        }

        // Проанализировать документ filePath и создать JSON файл в директории resultDirPath со свойствами колонтитулов типа type
        public void GenerateHeadersFootersInfoJSON(HeaderFooterType type, string filePath, string resultDirPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            List<HeaderFooterInfoGemBox> headersFootersInfo = PropertiesPuller.GetHeadersFootersInfo(type: type, filePath: filePath);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string headersFootersInfoJSON = JSONWorker.MakeJSON(headersFootersInfo);
            string resultFilePath = Path.Combine(resultDirPath, DefaultFileNames.HeadersFootersInfoFileName);
            FileWorker.WriteToFile(resultFilePath, headersFootersInfoJSON);
        }

        // Проанализировать документ filePath и создать csv файл resultPath со свойствами его параграфов (ID первого параграфа paragraphID)
        public void GenerateParagraphsPropertiesCSV(string filePath, int paragraphID, string resultPath, bool silent = false)
        {
            if (!silent) { Console.WriteLine($"Started {Path.GetFileName(filePath)}"); }
            List<ParagraphPropertiesGemBox> propertiesForFile = new List<ParagraphPropertiesGemBox>();
            string time = TimeCounter.GetExecutionTime(() => { propertiesForFile = PropertiesPuller.GetAllParagraphsProperties(filePath: filePath, paragraphID: paragraphID); }, TimeCounter.ResultType.TotalMilliseconds);
            if (!silent) { Console.WriteLine($"Done {Path.GetFileName(filePath)} in {time}"); }
            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.ParagraphsPropertiesFileName) : resultPath;
            FileWorker.FillCSV(resultFilePath, propertiesForFile);
        }

        // Пройтись по всем поддиректориям rootDir и в каждой создать csv файл, где будут записаны свойства параграфов для всех docx файлов в этой директории
        public void GenerateCSVFiles(string rootDir)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphPropertiesGemBox> propertiesForDir = new List<ParagraphPropertiesGemBox>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphPropertiesGemBox> propertiesForFile = PropertiesPuller.GetAllParagraphsProperties(filePath: filePath, paragraphID: 0);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                });

                string resultFilePath = Path.Combine(subDir, DefaultFileNames.ParagraphsPropertiesFileName);
                FileWorker.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // Проанализировать документ filePath и создать csv файл resultPath со свойствами его параграфов (ДЛЯ ТАБЛИЦЫ 0) (ID первого параграфа paragraphID)
        public void GenerateParagraphsPropertiesCSVForTableZero(string filePath, int paragraphID, string resultPath, bool silent = false)
        {
            if (!silent) { Console.WriteLine($"Started {Path.GetFileName(filePath)}"); }
            List<ParagraphPropertiesTableZero> propertiesForFile = new List<ParagraphPropertiesTableZero>();
            string time = TimeCounter.GetExecutionTime(() => { propertiesForFile = PropertiesPuller.GetAllParagraphsPropertiesForTableZero(filePath: filePath, paragraphID: paragraphID); }, TimeCounter.ResultType.TotalMilliseconds);
            if (!silent) { Console.WriteLine($"Done {Path.GetFileName(filePath)} in {time}"); }
            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.ParagraphsPropertiesForTableZeroFileName) : resultPath;
            FileWorker.FillCSV(resultFilePath, propertiesForFile);
        }

        // Запустить обычный GenerateParagraphProperties и GenerateParagraphsPropertiesCSVForTableZero
        public void GenerateParagraphsPropertiesForAllTables(string filePath, int paragraphID, string resultPath1, string resultPath2)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            Task firstTableTask = Task.Run(() => GenerateParagraphsPropertiesCSV(filePath, paragraphID, resultPath1, silent: true));
            Task secondTableTask = Task.Run(() => GenerateParagraphsPropertiesCSVForTableZero(filePath, paragraphID, resultPath2, silent: true));
            string time = TimeCounter.GetExecutionTime(() => { Task.WaitAll(firstTableTask, secondTableTask); }, TimeCounter.ResultType.TotalMilliseconds);
            Console.WriteLine($"Done {Path.GetFileName(filePath)} in {time}");
        }

        // GenerateCSVFiles, основанный на асинхронном методе
        public void GenerateCSVFilesAsync(string rootDir)
        {
            IPropertiesPullerAsync? asyncPuller = PropertiesPuller as IPropertiesPullerAsync;

            if (asyncPuller == null) { return; }

            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphPropertiesGemBox> propertiesForDir = new List<ParagraphPropertiesGemBox>();

                DirectoryIterator.IterateDocxFiles(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphPropertiesGemBox> propertiesForFile = asyncPuller.GetAllParagraphsPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                });

                string resultFilePath = Path.Combine(subDir, DefaultFileNames.AsyncParagraphsSyncIterationFileName);
                FileWorker.FillCSV(resultFilePath, propertiesForDir);
            });
        }

        // GenerateCSVFiles с асинхронным анализом файлов (ID первого параграфа paragraphID)
        public void GenerateCSVFilesWithAsyncFilesIteration(string rootDir)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphPropertiesGemBox> propertiesForDir = new List<ParagraphPropertiesGemBox>();

                Task.WaitAll(DirectoryIterator.IterateDocxFilesAsync(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphPropertiesGemBox> propertiesForFile = PropertiesPuller.GetAllParagraphsProperties(filePath: filePath, paragraphID: 0);
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                }));

                string resultFilePath = Path.Combine(subDir, DefaultFileNames.SyncParagraphsAsyncIterationFileName);
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
                List<ParagraphPropertiesGemBox> propertiesForDir = new List<ParagraphPropertiesGemBox>();

                Task.WaitAll(DirectoryIterator.IterateDocxFilesAsync(subDir, (filePath) =>
                {
                    Console.WriteLine($"Started {Path.GetFileName(filePath)}");
                    List<ParagraphPropertiesGemBox> propertiesForFile = asyncPuller.GetAllParagraphsPropertiesAsync(filePath: filePath).Result;
                    Console.WriteLine($"Done {Path.GetFileName(filePath)}");
                    propertiesForDir.AddRange(propertiesForFile);
                }));

                string resultFilePath = Path.Combine(subDir, DefaultFileNames.AsyncParagraphsAsyncIterationFileName);
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

        // Получить список ошибок форматирования для ВСЕГО документа filePath по требованиям (ГОСТу) rules с учетом классификации paragraphClasses
        // Сохранение результата по пути resultDirPath
        public void GenerateMistakesJSON(string fileToCorrect, RulesModel rules, string paragraphsClassesFile, string resultPath)
        {
            List<ClassificationResult>? paragraphsClassesList = JSONWorker.DeserializeObjectFromFile<List<ClassificationResult>>(paragraphsClassesFile);
            if (paragraphsClassesList == null) { return; }

            Console.WriteLine($"Started {Path.GetFileName(fileToCorrect)}");
            DocumentCorrections documentCorrections = new DocumentCorrections(rules);
            string time = TimeCounter.GetExecutionTime(() => { documentCorrections = Corrector.GetCorrections(fileToCorrect, rules, paragraphsClassesList); }, TimeCounter.ResultType.TotalMilliseconds);
            Console.WriteLine($"Done {Path.GetFileName(fileToCorrect)} in {time}");

            string documentCorrectionsJSON = JSONWorker.MakeJSON(documentCorrections);
            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.MistakesFileName) : resultPath;
            FileWorker.WriteToFile(resultFilePath, documentCorrectionsJSON);
        }

        // MARK: НИРМА 2020-2021
        // Получить файл, содержащий свойства параграфов документа filePath + проставить там возможные классы из файла с пресетами presetsPath (ID первого параграфа paragraphID)
        // Сохранение результата по пути resultPath
        public void GenerateCSVWithPresetsInfo(string filePath, string presetsPath, int paragraphID, string resultPath, bool silent = false)
        {
            CombinedPresetValues? combinedPresetValues = JSONWorker.DeserializeObjectFromFile<CombinedPresetValues>(presetsPath);
            if (combinedPresetValues == null) { return; }

            if (!silent) { Console.WriteLine($"Started {Path.GetFileName(filePath)}"); }
            List<ParagraphPropertiesWithPresets> propertiesForFile = new List<ParagraphPropertiesWithPresets>();
            string time = TimeCounter.GetExecutionTime(() => { propertiesForFile = PropertiesPuller.GetAllParagraphPropertiesWithPresets(filePath, paragraphID, combinedPresetValues); }, TimeCounter.ResultType.TotalMilliseconds);
            if (!silent) { Console.WriteLine($"Done {Path.GetFileName(filePath)} in {time}"); }

            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.ParagraphsPropertiesWithPresets) : resultPath;
            FileWorker.FillCSV(resultFilePath, propertiesForFile);
        }

        // Получить список ошибок форматирования одного абзаца документа fileToCorrect по требованиям (ГОСТу) rules (номер абзаца paragraphID, класс абзаца chosenClass)
        // Сохранение результата по пути resultDirPath
        public void GenerateFormattingMistakesJSON(string fileToCorrect, RulesModel rules, int paragraphID, ParagraphClass paragraphClass, string resultPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(fileToCorrect)}");
            ParagraphCorrections? paragraphCorrections = null;
            string time = TimeCounter.GetExecutionTime(() => { paragraphCorrections = Corrector.GetSingleParagraphCorrections(fileToCorrect, rules, paragraphID, paragraphClass); }, TimeCounter.ResultType.TotalMilliseconds);
            Console.WriteLine($"Done {Path.GetFileName(fileToCorrect)} in {time}");

            string paragraphCorrectionsJSON = JSONWorker.MakeJSON(paragraphCorrections);
            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.MistakesFileName) : resultPath;
            FileWorker.WriteToFile(resultFilePath, paragraphCorrectionsJSON);
        }

        // Получить файл, содержащий правила оформления класса paragraphClass для требований (ГОСТа) rules
        // Сохранение результата по пути resultPath
        public void GenerateModelJSON(RulesModel rules, ParagraphClass paragraphClass, string resultPath)
        {
            Console.WriteLine($"Started");
            PresetValue? classModel = PropertiesPuller.GetClassModel(rules, paragraphClass);
            Console.WriteLine($"Done");
            string classModelJSON = JSONWorker.MakeJSON(classModel);
            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.ClassModel) : resultPath;
            FileWorker.WriteToFile(resultFilePath, classModelJSON);
        }

        // Получить файл, содержащий структуру, которую можно использовать для примера значений пресета
        // Значения будут браться из параграфа под номером paragraphID документа filePath
        // Сохранение результата по пути resultPath
        public void GeneratePresetInfoFromParagraph(string filePath, int paragraphID, string resultPath)
        {
            Console.WriteLine($"Started {Path.GetFileName(filePath)}");
            PresetValue? paragraphPresetInfo = PropertiesPuller.GetParagraphPresetInfo(filePath, paragraphID);
            Console.WriteLine($"Done {Path.GetFileName(filePath)}");
            string paragraphPresetInfoJSON = JSONWorker.MakeJSON(paragraphPresetInfo);
            string resultFilePath = Directory.Exists(resultPath) ? Path.Combine(resultPath, DefaultFileNames.ParagraphPresetInfo) : resultPath;
            FileWorker.WriteToFile(resultFilePath, paragraphPresetInfoJSON);
        }

        // Пройтись по всем поддиректориям rootDir и получить ошибки для каждого абзаца находящегося там файла по требованиям (ГОСТу) rules
        // В каждой из поддиректорий должен находится один docx файл
        // В каждой из поддиректорий должен находится json файл с классами параграфов docx файла
        // В результате в каждой из поддиректорий будет создан файл с результатами проверки соответствующих параграфов
        public void GenerateParagraphMistakesFiles(string rootDir, RulesModel rules)
        {
            DirectoryIterator.IterateDir(rootDir, (subDir) =>
            {
                List<ParagraphCorrections> paragraphCorrectionsForDir = new List<ParagraphCorrections>();

                // Поиск нужных файлов в директории
                string docxFilePath;
                string resultFilePrefix;
                string classesJsonFilePath;
                string resultFilePath;
                try
                {
                    docxFilePath = Directory.EnumerateFiles(subDir, "*.docx", SearchOption.AllDirectories).FirstOrDefault();
                    resultFilePrefix = Path.GetFileNameWithoutExtension(docxFilePath);
                    classesJsonFilePath = Directory.EnumerateFiles(subDir, $"{resultFilePrefix}.json").FirstOrDefault();
                    string resultFileName = resultFilePrefix + "_mistakes.json";
                    resultFilePath = Path.Combine(subDir, resultFileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return;
                }

                List<ClassificationResult>? paragraphsClassesList = JSONWorker.DeserializeObjectFromFile<List<ClassificationResult>>(classesJsonFilePath);
                if (paragraphsClassesList == null) {
                    Console.WriteLine($"No classes json found in {Path.GetDirectoryName(docxFilePath)}");
                    return; 
                }

                List<ParagraphCorrections> directoryParagraphCorrections = new List<ParagraphCorrections>();

                Console.WriteLine($"Started {Path.GetFileName(docxFilePath)}");

                foreach (ClassificationResult classificationResult in paragraphsClassesList)
                {
                    ParagraphCorrections? paragraphCorrections = Corrector.GetSingleParagraphCorrections(docxFilePath, rules, classificationResult.Id, classificationResult.ParagraphClass);
                    if (paragraphCorrections != null) { directoryParagraphCorrections.Add(paragraphCorrections); }
                }

                Console.WriteLine($"Done {Path.GetFileName(docxFilePath)}");

                string paragraphCorrectionsJSON = JSONWorker.MakeJSON(directoryParagraphCorrections);
                FileWorker.WriteToFile(resultFilePath, paragraphCorrectionsJSON);
            });
        }
    }
}
using System;
using System.Collections.Generic;
using DocxCorrector.Services.Corrector;
using DocxCorrector.Services;
using DocxCorrector.Models;

namespace DocxCorrector.App
{
    class Program
    {
        public static Corrector Corrector = new CorrectorInterop();

        static void Main(string[] args)
        {
            // Выбирать документ
            Corrector.FilePath = Config.DocFilePath;

            // Проверить выбранный документ
            CheckParagraphs();

            Console.WriteLine("End of program");
            Console.ReadLine();
        }

        // Создать JSON файл с ошибками
        static void GenerateMistakesJSON()
        {
            List<ParagraphResult> mistakes = Corrector.GetMistakes();
            string mistakesJSON = JSONMaker.MakeJSON(mistakes);
            FileWriter.WriteToFile(Config.MistakesFilePath, mistakesJSON);
        }

        static void GeneratePagesPropertiesJSON()
        {
            List<PageProperties> pagesProperties = Corrector.GetAllPagesProperties();
            string pagesPropertiesJSON = JSONMaker.MakeJSON(pagesProperties);
            FileWriter.WriteToFile(Config.PagesPropertiesFilePath, pagesPropertiesJSON);
        }

        // Пройтись по всем поддиректориям Config.FilesToInpectDirectoryPath и в каждой создать csv файл, где будут результаты для всех docx файлов в этой директории
        static void GenerateCSVFiles()
        {
            DirectoryIterator.IterateDir(Config.FilesToInpectDirectoryPath, (subDir) =>
            {
                List<ParagraphProperties> propertiesForDir = new List<ParagraphProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filepath) =>
                {
                    Corrector.FilePath = filepath;
                    List<ParagraphProperties> propertiesForFile = Corrector.GetAllParagraphsProperties();
                    propertiesForFile.Add(new ParagraphProperties(paragraph: null));
                    propertiesForDir.AddRange(propertiesForFile);
                });

                FileWriter.FillPropertiesCSV(String.Concat(subDir, @"\results.csv"), propertiesForDir);
            });
        }

        // Получение данных для программы Ромы
        static void GenerateNormalizedCSVFiles()
        {
            DirectoryIterator.IterateDir(Config.FilesToInpectDirectoryPath, (subDir) =>
            {
                List<NormalizedProperties> normalizedPropertiesForDir = new List<NormalizedProperties>();

                DirectoryIterator.IterateDocxFiles(subDir, (filepath) =>
                {
                    Corrector.FilePath = filepath;
                    List<NormalizedProperties> normalizedPropertiesForFile = Corrector.GetNormalizedProperties();
                    normalizedPropertiesForDir.AddRange(normalizedPropertiesForFile);
                });

                FileWriter.FillPropertiesCSV(String.Concat(subDir, @"\normalizedResults.csv"), normalizedPropertiesForDir);
            });
        }

        // Получить JSON со списком ошибок для выбранного документа, с учетом того, что все параграфы в нем определенного типа
        static void CheckParagraphs()
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
                    paragraphResults = Corrector.GetMistakesForElementType(elementType: (ElementType)userAnserInt);
                    break;

                default:
                    Console.WriteLine("Ответ не поддерживается");
                    return;
            }

            if (paragraphResults == null) { return; }
            string resultJSON = JSONMaker.MakeJSON(results: paragraphResults);
            FileWriter.WriteToFile(Config.MistakesFilePath, resultJSON);
        }
    }
}

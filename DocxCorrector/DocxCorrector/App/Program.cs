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
            // Write your code here...

            // Получение данных для программы Ромы
            //GenerateNormalizedCSVFiles();
            GenerateCSVFiles();

            Console.WriteLine("End of program");
            Console.ReadLine();
        }

        // Создать JSON файл с ошибками
        static void GenerateMistacesJSON()
        {
            string mistakesJSON = Corrector.GetMistakesJSON();
            FileWriter.WriteToFile(Config.MistakesFilePath, mistakesJSON);
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
                    propertiesForFile.Add(new ParagraphProperties());
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
    }
}

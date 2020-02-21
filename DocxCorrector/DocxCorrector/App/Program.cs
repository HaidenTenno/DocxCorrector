using System;
using System.Collections.Generic;
using DocxCorrector.Services.Corrector;
using DocxCorrector.Services;
using DocxCorrector.Models;

namespace DocxCorrector.App
{
    class Program
    {
        public static Corrector Corrector = new CorrectorInterop(Config.DocFilePath);

        static void Main(string[] args)
        {
            // Получить ошибки для файла
            //string mistakesJSON = Corrector.GetMistakesJSON();
            //FileWriter.WriteToFile(Config.MistakesFilePath, mistakesJSON);

            // Получить свойства всех параграфов файла
            //List<ParagraphProperties> paragraphProperties = Corrector.GetAllParagraphsProperties();
            //FileWriter.FillPropertiesCSV(Config.PropertiesFilePath, paragraphProperties);

            // Пройтись по всем поддиректориям Config.FilesToInpectDirectoryPath и в каждой создать csv файл, где будут результаты для всех docx файлов в этой директории
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

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}

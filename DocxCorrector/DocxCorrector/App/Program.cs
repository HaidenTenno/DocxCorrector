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
            //string mistakesJSON = Corrector.GetMistakesJSON();
            //FileWriter.WriteToFile(Config.MistakesFilePath, mistakesJSON);
            
            List<ParagraphProperties> paragraphProperties = Corrector.GetAllParagraphsProperties();
            FileWriter.FillPropertiesCSV(Config.PropertiesFilePath, paragraphProperties);

            Console.ReadLine();
        }  
    }
}

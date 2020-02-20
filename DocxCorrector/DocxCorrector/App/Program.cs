using System;
using DocxCorrector.Services.Corrector;
using DocxCorrector.Services;

namespace DocxCorrector.App
{
    class Program
    {
        public static Corrector Corrector = new CorrectorInterop(Config.DocFilePath);

        static void Main(string[] args)
        {
            // Interop.Word version
            //string mistakesJSON = Corrector.GetMistakesJSON();
            //FileWriter.WriteToFile(Config.MistakesFilePath, mistakesJSON);
            //Corrector.PrintAllParagraphs();

            // TODO: - Remove
            System.Collections.Generic.List<Models.ParagraphProperties> list = new System.Collections.Generic.List<Models.ParagraphProperties>();
            var properties = new Models.ParagraphProperties
            {
                ID = 1,
                Text = "Test"
            };
            list.Add(properties);
            FileWriter.FillPropertiesCSV(Config.PropertiesFilePath, list);

            Console.ReadLine();
        }  
    }
}

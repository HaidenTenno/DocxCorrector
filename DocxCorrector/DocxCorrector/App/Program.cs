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

            //Console.WriteLine();
            Console.WriteLine("First paragraph properties:");
            
            Corrector.PrintFirstParagraphProperties();

            Console.ReadLine();
        }  
    }
}

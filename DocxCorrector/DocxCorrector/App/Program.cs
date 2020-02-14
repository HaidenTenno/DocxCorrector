using System;
using System.IO;
using System.Resources;
using DocxCorrector.Services.Corrector;
using Word = Microsoft.Office.Interop.Word;


namespace DocxCorrector.App
{
    class Program
    {
        public static Corrector Corrector = new CorrectorInterop(Config.FilePath);

        static void Main(string[] args)
        {
            // Interop.Word version
            Corrector.PrintAllParagraphs();
        }  
    }
}

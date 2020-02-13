using System;
using System.IO;
using System.Resources;
using DocxCorrector.Services;
using Word = Microsoft.Office.Interop.Word;


namespace DocxCorrector
{
    class Program
    {
        public static Corrector corrector = new CorrectorImpementation();

        static void Main(string[] args)
        {
            // TODO: Remove
            corrector.SayHi();

            string filePath = @"C:\Users\haide\Desktop\docxcorrector_sharp\DocxCorrector\DocxCorrector\DocxFiles\Doc1.docx";

            // Interop.Word version

            Word.Application app = new Word.Application();
            Word.Document document = app.Documents.Open(filePath);

            foreach (Word.Paragraph paragraph in document.Paragraphs)
            {
                Console.WriteLine(paragraph.Range.Text);
            }

            document.Close();

            // GemBox.Document version

        }
                
    }
}

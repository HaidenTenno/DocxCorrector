using System;

namespace DocxCorrector.Services.Corrector
{
    public abstract class Corrector
    {
        public string FilePath { get; }

        public Corrector(string filePath)
        {
            FilePath = filePath;
        }

        public abstract void PrintAllParagraphs();

    }
}

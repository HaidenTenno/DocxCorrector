using System;
using System.Text;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorGemBox : Corrector
    {
        public CorrectorGemBox(string filePath) : base(filePath) 
        {
            throw new NotImplementedException();
        }

        // Corrector
        public override string GetMistakesJSON()
        {
            throw new NotImplementedException();
        }

        public override void PrintAllParagraphs()
        {
            throw new NotImplementedException();
        }
        
        public override void PrintFirstParagraphProperties()
        {
            throw new NotImplementedException();
        }
    }
}

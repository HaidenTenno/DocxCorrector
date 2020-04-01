using System;
using System.Collections.Generic;
using Word = GemBox.Document;

namespace DocxCorrector.Models
{
    public sealed class SectionPropertiesGemBox : SectionProperties
    {
        public int SectionNumber { get; }

        public SectionPropertiesGemBox(Word.Section section, int sectionNumber)
        {
            SectionNumber = sectionNumber;
        }
    }
}

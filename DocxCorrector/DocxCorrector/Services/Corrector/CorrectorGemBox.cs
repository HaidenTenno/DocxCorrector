#nullable enable
using System;
using System.Collections.Generic;
using DocxCorrector.Models;

namespace DocxCorrector.Services.Corrector
{
    public sealed class CorrectorGemBox : Corrector
    {
        // Corrector
        public CorrectorGemBox(string? filePath = null) : base(filePath) { }

        // Получение JSON-а со списком ошибок
        public override List<ParagraphResult> GetMistakes()
        {
            throw new NotImplementedException();
        }

        // Получить свойства всех параграфов
        public override List<ParagraphProperties> GetAllParagraphsProperties()
        {
            throw new NotImplementedException();
        }

        //Получить свойства всех страниц
        public override List<PageProperties> GetAllPagesProperties()
        {
            throw new NotImplementedException();
        }

        // Получить нормализованные свойства параграфов (Для классификатора Ромы)
        public override List<NormalizedProperties> GetNormalizedProperties()
        {
            throw new NotImplementedException();
        }

        // Печать всех абзацев
        public override void PrintAllParagraphs()
        {
            throw new NotImplementedException();
        }

        // Получить списк ошибок для выбранного документа, с учетом того, что все параграфы в нем типа elementType
        public override List<ParagraphResult> GetMistakesForElementType(ElementType elementType)
        {
            throw new NotImplementedException();
        }
    }
}

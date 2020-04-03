using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.Corrector
{
    public interface ICorrecorAsync
    {
        // Для уверенности, что интерфейс реализуют только наследники Correcor
        public Corrector Corrector { get; }

        // Асинхронно получить свойства всех параграфов
        public Task<List<ParagraphProperties>> GetAllParagraphsPropertiesAsync(string filePath);

        // Асинхронно получить нормализованные свойства параграфов (Для классификатора Ромы)
        public Task<List<NormalizedProperties>> GetNormalizedPropertiesAsync(string filePath);
    }
}

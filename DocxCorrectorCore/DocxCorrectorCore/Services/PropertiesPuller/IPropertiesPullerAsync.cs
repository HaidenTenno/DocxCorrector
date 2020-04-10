using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.PropertiesPuller
{
    public interface IPropertiesPullerAsync
    {
        // Для уверенности, что интерфейс реализуют только наследники Correcor
        public PropertiesPuller PropertiesPuller { get; }

        // Асинхронно получить свойства всех параграфов
        public Task<List<ParagraphProperties>> GetAllParagraphsPropertiesAsync(string filePath);

        // Асинхронно получить нормализованные свойства параграфов
        public Task<List<NormalizedProperties>> GetNormalizedPropertiesAsync(string filePath);
    }
}

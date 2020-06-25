using System.Collections.Generic;
using System.Threading.Tasks;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public interface IPropertiesPullerAsync
    {
        // Для уверенности, что интерфейс реализуют только наследники PropertiesPuller
        public PropertiesPuller PropertiesPuller { get; }

        // Асинхронно получить свойства всех параграфов
        public Task<List<ParagraphPropertiesGemBox>> GetAllParagraphsPropertiesAsync(string filePath);
    }
}

using System.Collections.Generic;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.Corrector
{
    public abstract class Corrector
    {
        // Protected
        // Получить список ошибок форматирования ОТДЕЛЬНЫХ АБЗАЦЕВ для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        //protected abstract List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ParagraphClass> paragraphClasses);
        // TODO: More

        // Public
        // TODO: Implement
        // Получить список ошибок форматирования для ВСЕГО документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        //public abstract List<DocumentCorrections> GetCorrections(string filePath, RulesModel rulesModel, List<ParagraphClass> paragraphClasses);

        // Печать всех абзацев документа filePath
        public abstract void PrintAllParagraphs(string filePath);
    }
}

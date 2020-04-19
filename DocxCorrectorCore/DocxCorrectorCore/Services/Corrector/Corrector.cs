using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using DocxCorrectorCore.Models;

namespace DocxCorrectorCore.Services.Corrector
{
    public abstract class Corrector
    {
        // Private
        private async Task<List<ParagraphCorrections>> GetParagraphsCorrectionsAsync(string filePath, RulesModel rulesModel, List<ParagraphClass> paragraphsClasses)
        {
            return await Task.Run(() =>
            {
                Console.WriteLine("Beginning of paragraph formatting errors analysis");
                var result = GetParagraphsCorrections(filePath, rulesModel, paragraphsClasses);
                Console.WriteLine("Ending of paragraph formatting errors analysis");
                return result;
            });
        }

        // Protected
        // Получить список ошибок форматирования ОТДЕЛЬНЫХ АБЗАЦЕВ для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected abstract List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ParagraphClass> paragraphsClasses);
        // TODO: More

        // Public
        // Получить список ошибок форматирования для ВСЕГО документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        public virtual DocumentCorrections GetCorrections(string filePath, RulesModel rulesModel, List<ParagraphClass> paragraphsClasses)
        {
            var paragraphsCorrectionsTask = GetParagraphsCorrectionsAsync(filePath, rulesModel, paragraphsClasses);

            Task.WaitAll(paragraphsCorrectionsTask);

            DocumentCorrections documentCorrections = new DocumentCorrections(
                rules: rulesModel,
                paragraphsCorrections: paragraphsCorrectionsTask.Result
            );

            return documentCorrections;
        }

        // Печать всех абзацев документа filePath
        public abstract void PrintAllParagraphs(string filePath);
    }
}

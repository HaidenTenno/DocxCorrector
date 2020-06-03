using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    public abstract class Corrector
    {
        // Private
        private async Task<List<ParagraphCorrections>> GetParagraphsCorrectionsAsync(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphsClasses)
        {
            return await Task.Run(() =>
            {
                Console.WriteLine("Beginning of paragraph formatting errors analysis");
                var result = GetParagraphsCorrections(filePath, rulesModel, paragraphsClasses);
                Console.WriteLine("Ending of paragraph formatting errors analysis");
                return result;
            });
        }

        private async Task<List<SourcesListCorrections>> GetSourcesListCorrectionsAsync(string filePath, RulesModel rulesModel)
        {
            return await Task.Run(() =>
            {
                Console.WriteLine("Beginning of sources list errors analysis");
                var result = GetSourcesListCorrections(filePath, rulesModel);
                Console.WriteLine("Ending of sources list errors analysis");
                return result;
            });
        }

        // Protected
        // Получить список ошибок форматирования ОТДЕЛЬНЫХ АБЗАЦЕВ для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected abstract List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphsClasses);
        
        // Получить список ошибок оформления списка литературы для документа filePath по требованиям (ГОСТу) rulesModel
        protected abstract List<SourcesListCorrections> GetSourcesListCorrections(string filePath, RulesModel rulesModel);
        // TODO: More

        // Public
        // Получить список ошибок форматирования для ВСЕГО документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        public virtual DocumentCorrections GetCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphsClasses)
        {
            var paragraphsCorrectionsTask = GetParagraphsCorrectionsAsync(filePath, rulesModel, paragraphsClasses);
            var sourcesListCorrectionsTask = GetSourcesListCorrectionsAsync(filePath, rulesModel);

            Task.WaitAll(paragraphsCorrectionsTask);

            DocumentCorrections documentCorrections = new DocumentCorrections(
                rules: rulesModel,
                paragraphsCorrections: paragraphsCorrectionsTask.Result,
                sourcesListCorrections: sourcesListCorrectionsTask.Result
            );

            return documentCorrections;
        }

        // Печать всех абзацев документа filePath
        public abstract void PrintAllParagraphs(string filePath);
    }
}

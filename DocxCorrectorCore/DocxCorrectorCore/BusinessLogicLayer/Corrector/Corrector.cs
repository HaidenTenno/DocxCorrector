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
                Console.WriteLine("Beginning of paragraphs formatting errors analysis");
                var result = GetParagraphsCorrections(filePath, rulesModel, paragraphsClasses);
                Console.WriteLine("Ending of paragraphs formatting errors analysis");
                return result;
            });
        }

        private async Task<List<SourcesListCorrections>> GetSourcesListCorrectionsAsync(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            return await Task.Run(() =>
            {
                Console.WriteLine("Beginning of sources lists errors analysis");
                var result = GetSourcesListCorrections(filePath, rulesModel, paragraphClasses);
                Console.WriteLine("Ending of sources lists errors analysis");
                return result;
            });
        }

        private async Task<List<TableCorrections>> GetTableCorrectionsAsync(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            return await Task.Run(() =>
            {
                Console.WriteLine("Beginning of tables errors analysis");
                var result = GetTableCorrections(filePath, rulesModel, paragraphClasses);
                Console.WriteLine("Ending of tables errors analysis");
                return result;
            });
        }

        protected async Task<List<HeadlingCorrections>> GetHeadlingCorrectionsAsync(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            return await Task.Run(() =>
            {
                Console.WriteLine("Beginning of headlings errors analysis");
                var result = GetHeadlingCorrections(filePath, rulesModel, paragraphClasses);
                Console.WriteLine("Ending of headlings errors analysis");
                return result;
            });
        }

        // Protected
        // Получить список ошибок форматирования ОТДЕЛЬНЫХ АБЗАЦЕВ для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected abstract List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphsClasses);

        // Получить список ошибок оформления списка литературы для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected abstract List<SourcesListCorrections> GetSourcesListCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses);

        // Получить список ошибок оформления таблиц для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected abstract List<TableCorrections> GetTableCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses);

        // Получить список ошибок оформления заголовков для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected abstract List<HeadlingCorrections> GetHeadlingCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses);
        // TODO: More

        // Public
        // Получить список ошибок форматирования для ВСЕГО документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        public virtual DocumentCorrections GetCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphsClasses)
        {
            var paragraphsCorrectionsTask = GetParagraphsCorrectionsAsync(filePath, rulesModel, paragraphsClasses);
            var sourcesListCorrectionsTask = GetSourcesListCorrectionsAsync(filePath, rulesModel, paragraphsClasses);
            var tablesCorrectionsTask = GetTableCorrectionsAsync(filePath, rulesModel, paragraphsClasses);
            var headlingCorrectionsTask = GetHeadlingCorrectionsAsync(filePath, rulesModel, paragraphsClasses);

            Task.WaitAll(paragraphsCorrectionsTask);

            DocumentCorrections documentCorrections = new DocumentCorrections(
                rules: rulesModel,
                paragraphsCorrections: paragraphsCorrectionsTask.Result,
                sourcesListCorrections: sourcesListCorrectionsTask.Result,
                tablesCorrections: tablesCorrectionsTask.Result,
                headlingCorrections: headlingCorrectionsTask.Result
            );

            return documentCorrections;
        }

        // Печать всех абзацев документа filePath
        public abstract void PrintAllParagraphs(string filePath);
    }
}

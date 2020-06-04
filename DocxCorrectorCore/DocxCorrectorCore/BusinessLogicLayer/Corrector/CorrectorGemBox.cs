using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    public sealed class CorrectorGemBox : Corrector
    {
        // Private
        // Проверить соответсвтие списка классов и параграфов документа
        private bool ListsEquivalenceVerification(string filePath, List<ParagraphClass> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return false; }

            int currentParagraphClassIndex = 0;
            int totalElementsCount = 0;
            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                if (currentParagraphClassIndex >= paragraphClasses.Count) { break; }

                var paragraphs = section.GetChildElements(recursively: false, filterElements: Word.ElementType.Paragraph);
                totalElementsCount += paragraphs.Count();
                foreach (Word.Paragraph paragraph in paragraphs)
                {
                    if (currentParagraphClassIndex >= paragraphClasses.Count) { break; }

                    if (paragraph.Content.ToString().Trim() == "") { totalElementsCount--; continue; }
                    if (paragraph.ListFormat.IsList) { totalElementsCount--; continue; }

                    Console.WriteLine($"CLASS {paragraphClasses[currentParagraphClassIndex]}, PARAGRAPH {GemBoxHelper.GetParagraphPrefix(paragraph, 20)}");

                    currentParagraphClassIndex++;
                }
            }
            Console.WriteLine($"current index {currentParagraphClassIndex}, totalClassesListCount {paragraphClasses.Count()}, totalElementsCount = {totalElementsCount}");

            return (currentParagraphClassIndex == totalElementsCount);
        }

        // Public
        public CorrectorGemBox()
        {
            GemBoxHelper.SetLicense();
        }

        // Corrector
        // Protected
        // Получить список ошибок форматирования ОТДЕЛЬНЫХ АБЗАЦЕВ для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphCorrections>(); }

            if (paragraphClasses.Count() == 0) { return new List<ParagraphCorrections>(); }

            List<ParagraphCorrections> paragraphsCorrections = new List<ParagraphCorrections>();

            List<ClassifiedParagraph> classifiedParagraphs = GemBoxHelper.CombineParagraphsWithClassificationResult(document, paragraphClasses);

            // TODO: Model switch
            int currentParagraphIndex = 0;
            foreach (ClassifiedParagraph classifiedParagraph in classifiedParagraphs)
            {
                if (classifiedParagraph.ParagraphClass == null) 
                {
                    currentParagraphIndex++;
                    continue; 
                }

                // ПРОВЕРКА НАЧИНАЕТСЯ
                ParagraphCorrections? currentParagraphCorrections = null;
                switch (classifiedParagraph.ParagraphClass)
                {
                    case ParagraphClass.c1:
                        var standardParagraph = new ParagraphRegular();
                        currentParagraphCorrections = standardParagraph.CheckFormatting(currentParagraphIndex, classifiedParagraphs);
                        break;
                    default:
                        break;
                }
                if (currentParagraphCorrections != null) { paragraphsCorrections.Add(currentParagraphCorrections); }

                currentParagraphIndex++;
            }

            return paragraphsCorrections;
        }

        // Получить список ошибок оформления списка литературы для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<SourcesListCorrections> GetSourcesListCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<SourcesListCorrections>(); }

            List<SourcesListCorrections> sourcesListCorrections = new List<SourcesListCorrections>();

            List<ClassifiedParagraph> classifiedParagraphs = GemBoxHelper.CombineParagraphsWithClassificationResult(document, paragraphClasses);

            // TODO: Model switch
            var standartParagraph = new SourcesListElement();

            // Идти по списку классифицированных элементов
            for (int classifiedParagraphIndex = 0; classifiedParagraphIndex < classifiedParagraphs.Count(); classifiedParagraphIndex++)
            {
                // Если класс не b1, то пропускаем
                if (classifiedParagraphs[classifiedParagraphIndex].ParagraphClass != ParagraphClass.b1) { continue; }

                // Если в параграфе нет ключевой фразы, то пропускаем
                if (!standartParagraph.KeyWords.Any(keyword => classifiedParagraphs[classifiedParagraphIndex].Paragraph.Content.ToString().Contains(keyword, StringComparison.OrdinalIgnoreCase))) { continue; }

                // Идем до конца документа ИЛИ пока не встретим следующий заголовок
                for (int sourcesListParagraphIndex = classifiedParagraphIndex + 1; sourcesListParagraphIndex < classifiedParagraphs.Count(); sourcesListParagraphIndex++)
                {
                    if (classifiedParagraphs[sourcesListParagraphIndex].ParagraphClass == null) { continue; }

                    if (classifiedParagraphs[sourcesListParagraphIndex].ParagraphClass == ParagraphClass.b1) { break; }

                    // ПРОВЕРКА НАЧИНАЕТСЯ
                    SourcesListCorrections? currentSourcesListCorrections = standartParagraph.CheckSourcesList(sourcesListParagraphIndex, classifiedParagraphs);

                    if (currentSourcesListCorrections != null) { sourcesListCorrections.Add(currentSourcesListCorrections); }
                }
            }

            return sourcesListCorrections;
        }

        // Public
        // Печать всех абзацев документа filePath
        public override void PrintAllParagraphs(string filePath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return; }

            foreach (Word.Paragraph paragraph in document.GetChildElements(recursively: true, filterElements: Word.ElementType.Paragraph))
            {
                int elementWitDifferentStyleCount = paragraph.GetChildElements(true, Word.ElementType.Run).Count();
                Console.WriteLine($"В этом параграфе {elementWitDifferentStyleCount} элемент(ов) с разным оформлением");
                foreach (Word.Run run in paragraph.GetChildElements(recursively: true, filterElements: Word.ElementType.Run)) 
                {
                    string text = run.Text;
                    Console.WriteLine(text);
                }
                Console.WriteLine();
            }
        }
    }
}

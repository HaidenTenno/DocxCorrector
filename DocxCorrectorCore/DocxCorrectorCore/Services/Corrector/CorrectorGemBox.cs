using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.Models;
using DocxCorrectorCore.Models.ElementsObjectModel;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Services.Corrector
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

            List<Word.Paragraph> paragraphs = new List<Word.Paragraph>();
            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (Word.Paragraph paragraph in section.GetChildElements(recursively: false, filterElements: Word.ElementType.Paragraph))
                {
                    paragraphs.Add(paragraph);
                }
            }

            // TODO: Model switch

            int currentClassIndex = 0;
            int currentParagraphIndex = 0;
            foreach (Word.Paragraph paragraph in paragraphs)
            {
                int currentParagraphClassIndex;
                try { currentParagraphClassIndex = paragraphClasses[currentClassIndex].Id; } catch { return paragraphsCorrections; }
                if (currentParagraphIndex < currentParagraphClassIndex)
                {
                    currentParagraphIndex++;
                    continue;
                }

                // ПРОВЕРКА НАЧИНАЕТСЯ
                ParagraphCorrections? currentParagraphCorrections = null;
                switch (paragraphClasses[currentClassIndex].ParagraphClass)
                {
                    case ParagraphClass.c1:
                        var standardParagraph = new ParagraphRegular();
                        currentParagraphCorrections = standardParagraph.CheckFormatting(currentParagraphIndex, paragraphs);
                        break;
                    default:
                        break;
                }
                if (currentParagraphCorrections != null) { paragraphsCorrections.Add(currentParagraphCorrections); }

                currentParagraphIndex++;
                currentClassIndex++;
            }

            return paragraphsCorrections;
        }

        protected override List<SourcesListCorrections> GetSourcesListCorrections(string filePath, RulesModel rulesModel)
        {
            return new List<SourcesListCorrections> { SourcesListCorrections.TestSourcesListCorrection };
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

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

            // Идти по списку классифицированных элементов
            for (int classifiedParagraphIndex = 0; classifiedParagraphIndex < classifiedParagraphs.Count(); classifiedParagraphIndex++)
            {
                // Если клас не определен, то пропускаем
                if (classifiedParagraphs[classifiedParagraphIndex].ParagraphClass == null) { continue; }

                // ПРОВЕРКА НАЧИНАЕТСЯ
                ParagraphCorrections? currentParagraphCorrections = null;
                DocumentElement? standardParagraph = null;

                ParagraphClass? paragraphClass = classifiedParagraphs[classifiedParagraphIndex].ParagraphClass;
                switch (paragraphClass)
                {
                    case ParagraphClass.c1:
                        standardParagraph = new ParagraphRegular();
                        break;
                    case ParagraphClass.c2:
                        standardParagraph = new ParagraphBeforeList();
                        break;
                    case ParagraphClass.c3:
                        standardParagraph = new ParagraphBeforeEquation();
                        break;
                    case ParagraphClass.b1:
                        standardParagraph = new HeadingFirstLevel();
                        break;
                    case ParagraphClass.b2:
                    case ParagraphClass.b3:
                    case ParagraphClass.b4:
                        standardParagraph = new HeadingOtherLevels((ParagraphClass)paragraphClass);
                        break;
                    case ParagraphClass.d1:
                        standardParagraph = new SimpleListFirstElement();
                        break;
                    case ParagraphClass.d2:
                        standardParagraph = new SimpleListMiddleElement();
                        break;
                    case ParagraphClass.d3:
                        standardParagraph = new SimpleListLastElement();
                        break;
                    case ParagraphClass.d4:
                        standardParagraph = new ComplexListFirstElement();
                        break;
                    case ParagraphClass.d5:
                        standardParagraph = new ComplexListMiddleElement();
                        break;
                    case ParagraphClass.d6:
                        standardParagraph = new ComplexListLastElement();
                        break;
                    case ParagraphClass.h1:
                        standardParagraph = new ImageSign();
                        break;
                    case ParagraphClass.f1:
                    case ParagraphClass.f3:
                    case ParagraphClass.f5:
                        standardParagraph = new TableSign((ParagraphClass)paragraphClass);
                        break;
                    case ParagraphClass.r0:
                        standardParagraph = new SourcesListElement();
                        break;
                    default:
                        break;
                }
                if (standardParagraph != null) { currentParagraphCorrections = standardParagraph.CheckFormatting(classifiedParagraphIndex, classifiedParagraphs); }
                if (currentParagraphCorrections != null) { paragraphsCorrections.Add(currentParagraphCorrections); }
            }

            return paragraphsCorrections;
        }

        // Получить список ошибок оформления списка литературы для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<SourcesListCorrections> GetSourcesListCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<SourcesListCorrections>(); }

            if (paragraphClasses.Count() == 0) { return new List<SourcesListCorrections>(); }

            List<SourcesListCorrections> sourcesListCorrections = new List<SourcesListCorrections>();

            List<ClassifiedParagraph> classifiedParagraphs = GemBoxHelper.CombineParagraphsWithClassificationResult(document, paragraphClasses);

            // TODO: Model switch
            var standartSourcesList = new SourcesList();

            // Идти по списку классифицированных элементов
            for (int classifiedParagraphIndex = 0; classifiedParagraphIndex < classifiedParagraphs.Count(); classifiedParagraphIndex++)
            {
                // Если класс не b1, то пропускаем
                if (classifiedParagraphs[classifiedParagraphIndex].ParagraphClass != ParagraphClass.b1) { continue; }

                // Если в параграфе нет ключевой фразы, то пропускаем
                if (!standartSourcesList.KeyWords.Any(keyword => classifiedParagraphs[classifiedParagraphIndex].Element.Content.ToString().Contains(keyword, StringComparison.OrdinalIgnoreCase))) { continue; }

                // ПРОВЕРКА НАЧИНАЕТСЯ
                SourcesListCorrections? currentSourcesListCorrections = standartSourcesList.CheckSourcesList(classifiedParagraphIndex, classifiedParagraphs);

                if (currentSourcesListCorrections != null) { sourcesListCorrections.Add(currentSourcesListCorrections); }
            }

            return sourcesListCorrections;
        }

        // Получить список ошибок оформления таблиц для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<TableCorrections> GetTableCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<TableCorrections>(); }

            if (paragraphClasses.Count() == 0) { return new List<TableCorrections>(); }

            List<TableCorrections> tableCorrections = new List<TableCorrections>();

            List<ClassifiedParagraph> classifiedParagraphs = GemBoxHelper.CombineParagraphsWithClassificationResult(document, paragraphClasses);

            // TODO: Model switch
            var standartTable = new Table();

            for (int classifiedParagraphIndex = 0; classifiedParagraphIndex < classifiedParagraphs.Count(); classifiedParagraphIndex++)
            {
                // Если класс не e0, то пропускаем
                if (classifiedParagraphs[classifiedParagraphIndex].ParagraphClass != ParagraphClass.e0) { continue; }

                // ПРОВЕРКА НАЧИНАЕТСЯ
                TableCorrections? currentTableCorrections = standartTable.CheckTable(classifiedParagraphIndex, classifiedParagraphs);

                if (currentTableCorrections != null) { tableCorrections.Add(currentTableCorrections); }
            }

            return tableCorrections;
        }

        // Получить список ошибок оформления заголовков для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<HeadlingCorrections> GetHeadlingCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<HeadlingCorrections>(); }

            if (paragraphClasses.Count() == 0) { return new List<HeadlingCorrections>(); }

            List<HeadlingCorrections> headlingCorrections = new List<HeadlingCorrections>();

            List<ClassifiedParagraph> classifiedParagraphs = GemBoxHelper.CombineParagraphsWithClassificationResult(document, paragraphClasses);

            // TODO: Model switch

            // TODO: ПРОДОЛЖИТЬ ТУТ

            return new List<HeadlingCorrections> { HeadlingCorrections.TestHeadlingCorrection };
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

        // MARK: НИРМА 2020-2021
        //public override FixedDocument GetFixedDocument(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphsClasses)
        //{
        //    Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
        //    if (document == null) { return new FixedDocument(null, "FAIL TO OPEN"); }

        //    // TODO: Model switch

        //    FixedDocument fixedDocument = new FixedDocument(document, "SUCCESS");
        //    return fixedDocument;
        //}
    }
}

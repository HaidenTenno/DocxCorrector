using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel;
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
        // Получить список ошибок форматирования абзацев для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ClassificationResult> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphCorrections>(); }

            if (paragraphClasses.Count() == 0) { return new List<ParagraphCorrections>(); }

            List<ParagraphCorrections> paragraphsCorrections = new List<ParagraphCorrections>();

            List<ClassifiedParagraph> classifiedParagraphs = GemBoxHelper.CombineParagraphsWithClassificationResult(document, paragraphClasses);

            // Модель
            GlobalDocumentModel model = ModelSwitcher.GetSelectedModel(rulesModel);

            // Идти по списку классифицированных элементов
            for (int classifiedParagraphIndex = 0; classifiedParagraphIndex < classifiedParagraphs.Count(); classifiedParagraphIndex++)
            {
                // Если клас не определен, то пропускаем
                if (classifiedParagraphs[classifiedParagraphIndex].ParagraphClass == null) { continue; }

                // ПРОВЕРКА НАЧИНАЕТСЯ
                ParagraphCorrections? currentParagraphCorrections = null;
                DocumentElement? standardParagraph = model.ParagraphFormattingModel.GetDocumentElementFromClass((ParagraphClass)classifiedParagraphs[classifiedParagraphIndex].ParagraphClass!);

                // TODO: Проверить, что в файле ParagraphFormattingModelGOST_7_32 отражены все классы ниже

                //ParagraphClass? paragraphClass = classifiedParagraphs[classifiedParagraphIndex].ParagraphClass;
                //switch (paragraphClass)
                //{
                //    case ParagraphClass.c1:
                //        standardParagraph = new ParagraphRegularGOST_7_32();
                //        break;
                //    case ParagraphClass.c2:
                //        standardParagraph = new ParagraphBeforeListGOST_7_32();
                //        break;
                //    case ParagraphClass.c3:
                //        standardParagraph = new ParagraphBeforeEquationGOST_7_32();
                //        break;
                //    case ParagraphClass.b1:
                //        standardParagraph = new HeadingFirstLevelGOST_7_32();
                //        break;
                //    case ParagraphClass.b2:
                //    case ParagraphClass.b3:
                //    case ParagraphClass.b4:
                //        standardParagraph = new HeadingOtherLevelsGOST_7_32((ParagraphClass)paragraphClass);
                //        break;
                //    case ParagraphClass.d1:
                //        standardParagraph = new SimpleListFirstElementGOST_7_32();
                //        break;
                //    case ParagraphClass.d2:
                //        standardParagraph = new SimpleListMiddleElementGOST_7_32();
                //        break;
                //    case ParagraphClass.d3:
                //        standardParagraph = new SimpleListLastElementGOST_7_32();
                //        break;
                //    case ParagraphClass.d4:
                //        standardParagraph = new ComplexListFirstElementGOST_7_32();
                //        break;
                //    case ParagraphClass.d5:
                //        standardParagraph = new ComplexListMiddleElementGOST_7_32();
                //        break;
                //    case ParagraphClass.d6:
                //        standardParagraph = new ComplexListLastElementGOST_7_32();
                //        break;
                //    case ParagraphClass.h1:
                //        standardParagraph = new ImageSignGOST_7_32();
                //        break;
                //    case ParagraphClass.f1:
                //    case ParagraphClass.f3:
                //    case ParagraphClass.f5:
                //        standardParagraph = new TableSignGOST_7_32((ParagraphClass)paragraphClass);
                //        break;
                //    case ParagraphClass.r0:
                //        standardParagraph = new SourcesListElementGOST_7_32();
                //        break;
                //    default:
                //        break;
                //}

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

            // TODO: Model switch + продолжить, когда будут разрабатываться новые режимы
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

            // TODO: Model switch + продолжить, когда будут разрабатываться новые режимы
            var standartTable = new TableGOST_7_32();

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

            // TODO: Model switch + продолжить, когда будут разрабатываться новые режимы

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
        // Получить список ошибок форматирования одного абзаца под номером paragraphID документа filePath по требованиям (ГОСТу) rulesModel с учетом класса paragraphClass
        public override ParagraphCorrections? GetSingleParagraphCorrections(string filePath, RulesModel rulesModel, int paragraphID, ParagraphClass paragraphClass)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return null; }

            if (paragraphID < 0) { Console.WriteLine("Paragraph ID must be a non-negative number"); return null; }

            // Модель
            GlobalDocumentModel model = ModelSwitcher.GetSelectedModel(rulesModel);

            // Берем все параграфы (это абзацы и таблицы)
            List<Word.Element> elements = new List<Word.Element>();
            foreach (Word.Section section in document.GetChildElements(recursively: false, filterElements: Word.ElementType.Section))
            {
                foreach (var element in section.GetChildElements(recursively: false, filterElements: new Word.ElementType[] { Word.ElementType.Paragraph, Word.ElementType.Table }))
                {
                    elements.Add(element);
                }
            }

            if (paragraphID >= elements.Count) { Console.WriteLine("The paragraph with the given id is not found"); return null; }

            // ПРОВЕРКА НАЧИНАЕТСЯ
            ParagraphCorrections? singleParagraphCorrections = null;
            DocumentElement? standardParagraph = model.ParagraphFormattingModel.GetDocumentElementFromClass(paragraphClass);
            Word.Element selectedElement = elements[paragraphID];

            // Проверка, что класс параграфа поддерживается
            if (standardParagraph == null)
            {
                Console.WriteLine($"Class {paragraphClass} is not supported right now");
                return ParagraphCorrections.NotSupportedParagraphCorrection(paragraphID, paragraphClass, GemBoxHelper.GetParagraphPrefix(selectedElement, 20));
            }

            // Получить нужный GemBox класс текущего параграфа
            if (selectedElement is Word.Paragraph paragraph)
            {
                singleParagraphCorrections = standardParagraph.CheckSingleParagraphFormatting(paragraphID, paragraph);
            }

            return singleParagraphCorrections;
        }
    }
}

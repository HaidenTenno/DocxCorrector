using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.Models;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Services.Corrector
{
    public sealed class CorrectorGemBox : Corrector
    {
        // Public
        public CorrectorGemBox()
        {
            GemBoxHelper.SetLicense();
        }

        // Corrector
        // Protected
        // Получить список ошибок форматирования ОТДЕЛЬНЫХ АБЗАЦЕВ для документа filePath по требованиям (ГОСТу) rulesModel с учетом классификации paragraphClasses
        protected override List<ParagraphCorrections> GetParagraphsCorrections(string filePath, RulesModel rulesModel, List<ParagraphClass> paragraphClasses)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return new List<ParagraphCorrections>(); }

            List<ParagraphCorrections> paragraphsCorrections = new List<ParagraphCorrections>();

            // TODO: NEXT STEP (AFTER ELEMENTS OBJECT MODEL IMPLEMENTATION)

            // TODO: REMOVE
            ParagraphCorrections testCorrection = new ParagraphCorrections(
                paragraphID: 0,
                paragraphClass: paragraphClasses[0],
                prefix: "Test prefix",
                mistakes: new List<ParagraphMistake>
                {
                    new ParagraphMistake(
                        message: "No mistake",
                        advice: "Nothing to advice"
                    )
                });
            paragraphsCorrections.Add(testCorrection);

            return paragraphsCorrections;
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

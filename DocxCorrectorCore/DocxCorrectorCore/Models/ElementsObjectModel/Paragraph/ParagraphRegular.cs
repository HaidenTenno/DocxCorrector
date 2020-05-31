using System;
using System.Linq;
using System.Collections.Generic;
using GemBox.Document;
using DocxCorrectorCore.Services.Helpers;


namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ParagraphRegular : DocumentElement
    {
        //c1

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c1;

        // Свойства ParagraphFormat
        public override HorizontalAlignment Alignment => HorizontalAlignment.Justify;
        public override bool KeepWithNext => false;
        public override OutlineLevel OutlineLevel => OutlineLevel.BodyText;
        public override bool PageBreakBefore => false;
        public override double SpecialIndentationLeftBorder => -36.85;
        public override double SpecialIndentationRightBorder => -35.45;

        // Свойства CharacterFormat для всего абзаца
        public override bool WholeParagraphAllCaps => false;
        public override bool WholeParagraphBold => false;
        public override bool WholeParagraphSmallCaps => false;
        
        // Свойства CharacterFormat для всего абзаца
        public override bool RunnerBold => false;
        
        // Особые свойства
        //public override StartSymbolType? StartSymbol => StartSymbolType.Upper;
        public override int EmptyLinesAfter => 0;

        // TODO: Сделать проверку окончания абзаца регуляркой (а нужна ли она вообще?..) 
        // public override string[] Suffixes => new string[] { ".", "!", "?" }; - могут пройти ".." и др.

        // Метод проверки
        public override ParagraphCorrections? CheckFormatting(int id, List<Paragraph> paragraphs)
        {
            Paragraph paragraph;
            try { paragraph = paragraphs[id]; } catch { return null; }

            ParagraphCorrections? result = base.CheckFormatting(id, paragraphs);
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Особые свойства
            if ((paragraph.Content.ToString().Count() > 0) & (!char.IsUpper(paragraph.Content.ToString()[0])))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: "Параграф должен начинаться с большой буквы",
                    advice: "ТУТ БУДЕТ СОВЕТ"
                );
                paragraphMistakes.Add(mistake);
            }

            // Prefixes?
            // Suffixes?
            
            // if ((paragraph != EmptyLinesAfter)
            // {
            //     ParagraphMistake mistake = new ParagraphMistake(
            //         message: "После параграфа не должно быть пустых строк",
            //         advice: "ТУТ БУДЕТ СОВЕТ"
            //     );
            //     paragraphMistakes.Add(mistake);
            // }

            if (paragraphMistakes.Count != 0)
            {
                if (result != null) 
                { 
                    result.Mistakes.AddRange(paragraphMistakes); 
                }
                else
                {
                    result = new ParagraphCorrections(
                        paragraphID: id,
                        paragraphClass: ParagraphClass,
                        prefix: GemBoxHelper.GetParagraphPrefix(paragraph, 20),
                        mistakes: paragraphMistakes
                    );
                }
            }

            return result;
        }
    }
}
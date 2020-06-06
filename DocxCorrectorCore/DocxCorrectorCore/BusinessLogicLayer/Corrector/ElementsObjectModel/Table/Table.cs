using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class Table : DocumentElement
    {
        //e0

        public override ParagraphClass ParagraphClass => ParagraphClass.e0;

        // Свойства ParagraphFormat
        public override Word.HorizontalAlignment Alignment => Word.HorizontalAlignment.Left;
        public override bool KeepWithNext => false;
        public override Word.OutlineLevel OutlineLevel => Word.OutlineLevel.BodyText;
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

        public override int EmptyLinesAfter => 0;

        // TODO: Предыдущий элемент - подпись к таблице (f-элементы) или начало/продолжение таблицы (e1/e2)

        // Табличные свойства


        public TableCorrections? CheckTable(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Tables.Table table;
            try { table = (Word.Tables.Table)classifiedParagraphs[id].Element; } catch { return null; }

            System.Console.WriteLine("!!!TABLE!!!");

            return TableCorrections.TestTableCorrection;
        }
    }
}
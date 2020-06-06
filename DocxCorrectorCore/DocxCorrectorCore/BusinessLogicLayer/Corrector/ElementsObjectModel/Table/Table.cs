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

        // ТАБЛИЧНЫЕ СВОЙСТВА
        // Свойства TableFormat


        // Свойства TableRowFormat


        // Свойства TableCellFormat


        public TableCorrections? CheckTable(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Tables.Table table;
            try { table = (Word.Tables.Table)classifiedParagraphs[id].Element; } catch { return null; }

            List<TableMistake> tableMistakes = new List<TableMistake>();

            // Свойства TableFormat


            // Свойства TableRowFormat
            foreach (Word.Tables.TableRow tableRow in table.Rows)
            {
                    

                // Свойства TableCellFormat
                foreach (Word.Tables.TableCell tableCell in tableRow.Cells)
                {
                    


                    foreach (Word.Paragraph paragraph in tableCell.GetChildElements(false, Word.ElementType.Paragraph))
                    {
                        // Свойства ParagraphFormat для абзаца внутри ячейки таблицы


                        // Свойства CharacterFormat для всего абзаца внутри ячейки таблицы


                        // Свойства CharacterFormat для раннеров внутри ячейки таблицы
                        foreach (Word.Run runner in paragraph.GetChildElements(false, Word.ElementType.Run))
                        {

                        }
                    }
                }
            }

            if (tableMistakes.Count != 0)
            {
                return new TableCorrections(
                    paragraphID: id,
                    mistakes: tableMistakes
                );
            }
            else
            {
                return null;
            }
        }
    }
}
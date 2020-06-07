using System;
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

        private List<TableMistake> CheckTableFormat(int id, Word.Tables.Table table)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            //foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            //{
            //    Console.WriteLine($"{borderType} === {table.TableFormat.Borders[borderType].Style}");//(table.TableFormat.Borders[borderType].Style);
            //}

            return tableMistakes;
        }

        private List<TableMistake> CheckTableRowFormat(int id, int tableRowIndex, Word.Tables.TableRow tableRow)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (tableRowIndex == 0)
            {
                // Проверка первой строки таблицы
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckTableCellFormat(int id, int tableRowIndex, int tableCellIndex, Word.Tables.TableCell tableCell)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (tableRowIndex == 0)
            {
                // Проверка первой строки таблицы
            }

            if (tableCellIndex == 0)
            {
                // Проверка ячейки в первом столбце таблицы
            }


            //foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            //{
            //    Console.WriteLine($"{borderType} === {tableCell.CellFormat.Borders[borderType].Style}");//(tableCell.CellFormat.Borders[borderType].Style);
            //}

            return tableMistakes;
        }

        private List<TableMistake> CheckParagraphFormat(int id, Word.Paragraph paragraph)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            //foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            //{
            //    Console.WriteLine($"{borderType} === {paragraph.ParagraphFormat.Borders[borderType].Style}");//(paragraph.ParagraphFormat.Borders[borderType].Style);
            //}

            return tableMistakes;
        }

        private List<TableMistake> CheckParagraphCharacterFormat(int id, Word.Paragraph paragraph)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();



            return tableMistakes;
        }

        private List<TableMistake> CheckRunnerCharacterFormat(int id, Word.Run run)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            //foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            //{
            //    Console.WriteLine($"{borderType} === {runner.CharacterFormat.Border.Style}");//(runner.CharacterFormat.Border.Style);
            //}

            return tableMistakes;
        }

        public TableCorrections? CheckTable(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Tables.Table table;
            try { table = (Word.Tables.Table)classifiedParagraphs[id].Element; } catch { return null; }

            List<TableMistake> tableMistakes = new List<TableMistake>();

            // Свойства TableFormat
            tableMistakes.AddRange(CheckTableFormat(id, table));

            for (int tableRowIndex = 0; tableRowIndex < table.Rows.Count; tableRowIndex++)
            {
                Word.Tables.TableRow tableRow = table.Rows[tableRowIndex];

                // Свойства TableRowFormat
                tableMistakes.AddRange(CheckTableRowFormat(id, tableRowIndex, tableRow));

                for (int tableCellIndex = 0; tableCellIndex < tableRow.Cells.Count; tableCellIndex++) 
                {
                    Word.Tables.TableCell tableCell = tableRow.Cells[tableCellIndex];

                    // Свойства TableCellFormat
                    tableMistakes.AddRange(CheckTableCellFormat(id, tableRowIndex, tableCellIndex, tableCell));

                    foreach (Word.Paragraph paragraph in tableCell.GetChildElements(false, Word.ElementType.Paragraph))
                    {
                        // Свойства ParagraphFormat для абзаца внутри ячейки таблицы
                        tableMistakes.AddRange(CheckParagraphFormat(id, paragraph));

                        // Свойства CharacterFormat для всего абзаца внутри ячейки таблицы
                        tableMistakes.AddRange(CheckParagraphCharacterFormat(id, paragraph));

                        foreach (Word.Run runner in paragraph.GetChildElements(false, Word.ElementType.Run))
                        {
                            // Свойства CharacterFormat для раннеров внутри ячейки таблицы
                            tableMistakes.AddRange(CheckRunnerCharacterFormat(id, runner));
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
using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class Table : DocumentElement
    {
        //e0

        public override ParagraphClass ParagraphClass => ParagraphClass.e0;

        // Свойства ParagraphFormat

        // Свойства CharacterFormat для всего абзаца
        public override double WholeParagraphSizeLeftBorder => 10;
        public static new double? WholeParagraphChosenSize { get; protected set; } = null;

        // Свойства CharacterFormat для раннеров

        // Особые свойства

        // Свойства TableFormat
        public virtual Word.HorizontalAlignment TableFormatFirstRowAlignment => Word.HorizontalAlignment.Center;
        public virtual Word.HorizontalAlignment TableFormatFirstColumnAlignment => Word.HorizontalAlignment.Justify;
        public virtual List<Word.Color> TableFormatBackgroundColors => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual Word.BorderStyle TableFormatOuterBorders => Word.BorderStyle.Single;
        public virtual Word.BorderStyle TableFormatDiagonalBorders => Word.BorderStyle.None;
        public virtual List<Word.BorderStyle> TableFormatAvailableInnerBorders => new List<Word.BorderStyle> { Word.BorderStyle.None, Word.BorderStyle.Single };
        public virtual int TableFormatColumnBandSize => 1;
        public virtual double TableFormatDefaultCellSpacing => 0;
        public virtual double TableFormatIndentFromLeft => 0;

        // Свойства TableRowFormat


        // Свойства TableCellFormat


        // Проверка границ
        private bool CheckTableFormatBorder(Word.Tables.Table table)
        {
            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                
            }

            return true;
        }

        private bool CheckParagraphFormatBorder(Word.Paragraph paragraph)
        {
            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                if (paragraph.ParagraphFormat.Borders[borderType].Style != BorderStyle)
                {
                    return false;
                }
            }
            return true;
        }

        private bool CheckWholeParagraphBorder(Word.Paragraph paragraph)
        {
            if (paragraph.CharacterFormatForParagraphMark.Border != WholeParagraphBorder)
            {
                return false;
            }
            return true;
        }

        private bool CheckRunnerBorder(Word.Run runner)
        {
            if (runner.CharacterFormat.Border != WholeParagraphBorder)
            {
                return false;
            }
            return true;
        }

        private TableMistake? CheckEmptyLines(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            // Проверка, что пустых строк достаточно
            int emptyLinesCount = 1;
            while ((emptyLinesCount <= EmptyLinesAfter) & (id + emptyLinesCount < classifiedParagraphs.Count))
            {
                int idToCheckEmpty = id + emptyLinesCount;

                Word.Paragraph paragraphToCheckForEmpty;
                // Если следующий элемент не параграф, то он не пустой
                try { paragraphToCheckForEmpty = (Word.Paragraph)classifiedParagraphs[idToCheckEmpty].Element; }
                catch
                {
                    return new TableMistake(
                        row: -1,
                        column: -1,
                        message: $"Неверное количество пропущенных параграфов (недостаточно)",
                        advice: $"ТУТ БУДЕТ СОВЕТ"
                    );
                }

                if (GemBoxHelper.GetParagraphContentWithoutNewLine(paragraphToCheckForEmpty) != "")
                {
                    return new TableMistake(
                        row: -1,
                        column: -1,
                        message: $"Неверное количество пропущенных параграфов (недостаточно)",
                        advice: $"ТУТ БУДЕТ СОВЕТ"
                    );
                }

                emptyLinesCount++;
            }

            // Проверка, что пустых строк не слишком много (проверка, что id + emptyLinesCount параграф не пустой)
            int idToCheckNotEmpty = id + emptyLinesCount;

            Word.Paragraph paragraphToCheckForNotEmpty;
            // Если следующий элемент не параграф, то он не пустой
            try { paragraphToCheckForNotEmpty = (Word.Paragraph)classifiedParagraphs[idToCheckNotEmpty].Element; } catch { return null; }

            if (GemBoxHelper.GetParagraphContentWithoutNewLine(paragraphToCheckForNotEmpty) == "")
            {
                return new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Неверное количество пропущенных параграфов (слишком много)",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
            }

            return null;
        }

        private List<TableMistake> CheckTableFormat(Word.Tables.Table table)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            //Console.WriteLine(table.TableFormat.Positioning.DistanceFromSurroundingText);
            //Console.WriteLine(new Word.Padding(0,0,0,0));
            //Console.WriteLine();

            //Console.WriteLine(table.TableFormat.Positioning.HorizontalPosition);
            //Console.WriteLine(new Word.HorizontalPosition(Word.HorizontalPositionType.Center, Word.HorizontalPositionAnchor.Margin));
            //Console.WriteLine();

            //Console.WriteLine(table.TableFormat.Positioning.VerticalPosition);
            //Console.WriteLine(new Word.VerticalPosition(0, Word.LengthUnit.Centimeter, Word.VerticalPositionAnchor.Paragraph));
            //Console.WriteLine();

            return tableMistakes;
        }

        private List<TableMistake> CheckTableRowFormat(int tableRowIndex, Word.Tables.TableRow tableRow)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (tableRowIndex == 0)
            {
                // Проверка первой строки таблицы
            }


            return tableMistakes;
        }

        private List<TableMistake> CheckTableCellFormat(int tableRowIndex, int tableCellIndex, Word.Tables.TableCell tableCell)
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



            return tableMistakes;
        }

        private List<TableMistake> CheckParagraphFormat(int tableRowIndex, int tableCellIndex, Word.Paragraph paragraph)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!BackgroundColors.Contains(paragraph.ParagraphFormat.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверный цвет заливки параграфа",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckParagraphFormatBorder(paragraph))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"У параграфа присутствуют рамки",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.RightToLeft != RightToLeft)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Справа-налево'",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckParagraphCharacterFormat(int tableRowIndex, int tableCellIndex, Word.Paragraph paragraph)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!WholeParagraphBackgroundColors.Contains(paragraph.CharacterFormatForParagraphMark.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет заливки' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckWholeParagraphBorder(paragraph))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"У параграфа присутствуют рамки",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough != WholeParagraphDoubleStrikethrough)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.FontColor != WholeParagraphFontColor)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет шрифта' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.FontName != WholeParagraphFontName)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Шрифт' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Hidden != WholeParagraphHidden)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Скрытый' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphHighlightColors.Contains(paragraph.CharacterFormatForParagraphMark.HighlightColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет выделения' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Kerning != WholeParagraphKerning)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Кернинг' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.RightToLeft != WholeParagraphRightToLeft)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Справа-налево' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Scaling != WholeParagraphScaling)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Масштаб' для всего абзаца",
                    advice: "ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (paragraph.CharacterFormatForParagraphMark.Size != WholeParagraphChosenSize))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца (должно быть единообразие)",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }
            else
            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = paragraph.CharacterFormatForParagraphMark.Size;
            }

            if (paragraph.CharacterFormatForParagraphMark.Strikethrough != WholeParagraphStrikethrough)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Зачеркнутый' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Subscript != WholeParagraphSubscript)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Подстрочный' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Superscript != WholeParagraphSuperscript)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Надстрочный' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.UnderlineStyle != WholeParagraphUnderlineStyle)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Подчеркнутый' для всего абзаца",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckRunnerCharacterFormat(int tableRowIndex, int tableCellIndex, Word.Run runner)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!RunnerBackgroundColors.Contains(runner.CharacterFormat.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет заливки' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckRunnerBorder(runner))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Границы' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.DoubleStrikethrough != RunnerDoubleStrikethrough)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.FontColor != RunnerFontColor)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет шрифта' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.FontName != RunnerFontName)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Шрифт' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Hidden != RunnerHidden)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Скрытый' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerHighlightColors.Contains(runner.CharacterFormat.HighlightColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет выделения' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Kerning != RunnerKerning)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Кернинг' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.RightToLeft != RunnerRightToLeft)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Справа-налево' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Scaling != RunnerScaling)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Масштаб' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (runner.CharacterFormat.Size != WholeParagraphChosenSize))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера (должно быть единообразие)",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }
            else
            if ((runner.CharacterFormat.Size < RunnerSizeLeftBorder) | (runner.CharacterFormat.Size > RunnerSizeRightBorder))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = runner.CharacterFormat.Size;
            }

            if (runner.CharacterFormat.Strikethrough != RunnerStrikethrough)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Зачеркнутый' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.UnderlineStyle != RunnerUnderlineStyle)
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Подчеркнутый' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        public TableCorrections? CheckTable(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Tables.Table table;
            try { table = (Word.Tables.Table)classifiedParagraphs[id].Element; } catch { return null; }

            List<TableMistake> tableMistakes = new List<TableMistake>();

            // Свойства TableFormat
            tableMistakes.AddRange(CheckTableFormat(table));

            for (int tableRowIndex = 0; tableRowIndex < table.Rows.Count; tableRowIndex++)
            {
                Word.Tables.TableRow tableRow = table.Rows[tableRowIndex];

                // Свойства TableRowFormat
                tableMistakes.AddRange(CheckTableRowFormat(tableRowIndex, tableRow));

                for (int tableCellIndex = 0; tableCellIndex < tableRow.Cells.Count; tableCellIndex++) 
                {
                    Word.Tables.TableCell tableCell = tableRow.Cells[tableCellIndex];

                    // Свойства TableCellFormat
                    tableMistakes.AddRange(CheckTableCellFormat(tableRowIndex, tableCellIndex, tableCell));

                    foreach (Word.Paragraph paragraph in tableCell.GetChildElements(false, Word.ElementType.Paragraph))
                    {
                        // Свойства ParagraphFormat для абзаца внутри ячейки таблицы
                        tableMistakes.AddRange(CheckParagraphFormat(tableRowIndex, tableCellIndex, paragraph));

                        // Свойства CharacterFormat для всего абзаца внутри ячейки таблицы
                        tableMistakes.AddRange(CheckParagraphCharacterFormat(tableRowIndex, tableCellIndex, paragraph));

                        foreach (Word.Run runner in paragraph.GetChildElements(false, Word.ElementType.Run))
                        {
                            // Свойства CharacterFormat для раннеров внутри ячейки таблицы
                            tableMistakes.AddRange(CheckRunnerCharacterFormat(tableRowIndex, tableCellIndex, runner));
                        }
                    }
                }
            }

            // Особые свойства
            // Проверка количества пустых строк
            TableMistake? emptyLinesMistake = CheckEmptyLines(id, classifiedParagraphs);
            if (emptyLinesMistake != null) { tableMistakes.Add(emptyLinesMistake); }

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
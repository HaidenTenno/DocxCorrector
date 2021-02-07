using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class TableGOST_7_32 : DocumentElementGOST_7_32
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
        public virtual List<Word.HorizontalAlignment> FirstRowAlignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Center };
        public virtual List<Word.HorizontalAlignment> FirstColumnAlignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Justify };
        public virtual List<Word.Color> TableFormatBackgroundColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual List<Word.BorderStyle> TableFormatOuterBorders => new List<Word.BorderStyle> { Word.BorderStyle.Single };
        public virtual List<Word.BorderStyle> TableFormatDiagonalBorders => new List<Word.BorderStyle> { Word.BorderStyle.None };
        public virtual List<Word.BorderStyle> TableFormatAvailableInnerBorders => new List<Word.BorderStyle> { Word.BorderStyle.None, Word.BorderStyle.Single };
        public virtual List<int> TableFormatColumnBandSize => new List<int> { 1 };
        public virtual List<double> TableFormatDefaultCellSpacing => new List<double> { 0 };
        public virtual List<double> TableFormatIndentFromLeft => new List<double> { 0 };
        public virtual List<Word.Padding> TableFormatDistanceFromSurroundingText => new List<Word.Padding> { new Word.Padding(0, 0, 0, 0) };
        public virtual List<Word.HorizontalPosition> TableFormatHorizontalPosition => new List<Word.HorizontalPosition> { new Word.HorizontalPosition(Word.HorizontalPositionType.Absolute, Word.HorizontalPositionAnchor.Margin) };
        public virtual List<Word.VerticalPosition> TableFormatVerticalPosition => new List<Word.VerticalPosition> { new Word.VerticalPosition(0, Word.LengthUnit.Centimeter, Word.VerticalPositionAnchor.Margin) };
        public virtual List<bool> TableFormatRightToLeft => new List<bool> { false };
        public virtual List<int> TableFormatRowBandSize => new List<int> { 1 };

        // Свойства TableRowFormat
        public virtual List<bool> TableRowFormatAllowBreakAcrossPages => new List<bool> { true };
        public virtual List<bool> TableRowFormatHidden => new List<bool> { false };

        // Свойства TableCellFormat
        public virtual List<Word.Color> TableCellFormatBackgroundColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual List<Word.BorderStyle> TableCellFormatAvailableBorders => new List<Word.BorderStyle> { Word.BorderStyle.None, Word.BorderStyle.Single };
        public virtual List<Word.Tables.TableCellTextDirection> TableCellFormatTextDirection => new List<Word.Tables.TableCellTextDirection> { Word.Tables.TableCellTextDirection.LeftToRight };

        // Проверка границ
        private bool CheckTableFormatBorder(Word.Tables.Table table)
        {
            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                switch (borderType)
                {
                    case Word.SingleBorderType.Top:
                    case Word.SingleBorderType.Bottom:
                    case Word.SingleBorderType.Left:
                    case Word.SingleBorderType.Right:
                        if (!TableFormatOuterBorders.Contains(table.TableFormat.Borders[borderType].Style)) { return false; }
                        break;

                    case Word.SingleBorderType.InsideVertical:
                    case Word.SingleBorderType.InsideHorizontal:
                        if (!TableFormatAvailableInnerBorders.Contains(table.TableFormat.Borders[borderType].Style)) { return false; }
                        break;

                    case Word.SingleBorderType.DiagonalDown:
                    case Word.SingleBorderType.DiagonalUp:
                        if (!TableFormatDiagonalBorders.Contains(table.TableFormat.Borders[borderType].Style)) { return false; }
                        break;
                }
            }

            return true;
        }

        private bool CheckTableCellFormatBorder(Word.Tables.TableCell cell)
        {
            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                if (!TableCellFormatAvailableBorders.Contains(cell.CellFormat.Borders[borderType].Style)) { return false; }
            }

            return true;
        }

        private bool CheckParagraphFormatBorder(Word.Paragraph paragraph)
        {
            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                if (!BorderStyle.Contains(paragraph.ParagraphFormat.Borders[borderType].Style))
                {
                    return false;
                }
            }
            return true;
        }

        private bool CheckWholeParagraphBorder(Word.Paragraph paragraph)
        {
            if (!WholeParagraphBorder.Contains(paragraph.CharacterFormatForParagraphMark.Border))
            {
                return false;
            }
            return true;
        }

        private bool CheckRunnerBorder(Word.Run runner)
        {
            if (!WholeParagraphBorder.Contains(runner.CharacterFormat.Border))
            {
                return false;
            }
            return true;
        }

        private TableMistake? CheckEmptyLines(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            // Посчитать количество строк до следующего параграфа
            int emptyLinesCount = 1;
            while (id + emptyLinesCount < classifiedParagraphs.Count)
            {
                int idToCheckEmpty = id + emptyLinesCount;

                Word.Paragraph paragraphToCheckForEmpty;
                // Если следующий элемент не параграф, то он не пустой
                try { paragraphToCheckForEmpty = (Word.Paragraph)classifiedParagraphs[idToCheckEmpty].Element; }
                catch { break; }

                if (GemBoxHelper.GetParagraphContentWithoutNewLine(paragraphToCheckForEmpty) == "") { break; }

                emptyLinesCount++;
            }

            if (!EmptyLinesAfter.Contains(emptyLinesCount - 1))
            {
                return new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Неверное количество пропущенных параграфов"
                );
            }

            return null;
        }

        private List<TableMistake> CheckTableFormat(Word.Tables.Table table)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            // Проверка Alignment в других методах

            if (!TableFormatBackgroundColor.Contains(table.TableFormat.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Неверный цвет заливки таблицы"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckTableFormatBorder(table))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в рамках таблицы"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatColumnBandSize.Contains(table.TableFormat.ColumnBandSize))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в объединении столбцов"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatDefaultCellSpacing.Contains(table.TableFormat.DefaultCellSpacing))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в интервале между ячейками по умолчанию"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatIndentFromLeft.Contains(table.TableFormat.IndentFromLeft))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в отступе слева"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatDistanceFromSurroundingText.Contains(table.TableFormat.Positioning.DistanceFromSurroundingText))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в расстоянии до текста при обтекании"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatHorizontalPosition.Contains(table.TableFormat.Positioning.HorizontalPosition))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в горизонтальном положении при обтекании текстом"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatVerticalPosition.Contains(table.TableFormat.Positioning.VerticalPosition))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в вертикальном положении при обтекании текстом"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatRightToLeft.Contains(table.TableFormat.RightToLeft))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в свойстве 'слева-направо' для таблиц"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableFormatRowBandSize.Contains(table.TableFormat.RowBandSize))
            {
                TableMistake mistake = new TableMistake(
                    row: -1,
                    column: -1,
                    message: $"Ошибка в объединении строк"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckTableRowFormat(int tableRowIndex, Word.Tables.TableRow tableRow)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!TableRowFormatAllowBreakAcrossPages.Contains(tableRow.RowFormat.AllowBreakAcrossPages))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: -1,
                    message: $"Ошибка в свойстве 'Разрешить перенос строк на следующую страницу'"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableRowFormatHidden.Contains(tableRow.RowFormat.Hidden))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: -1,
                    message: $"Ошибка в свойстве 'скрытый'"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckTableCellFormat(int tableRowIndex, int tableCellIndex, Word.Tables.TableCell tableCell)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!TableCellFormatBackgroundColor.Contains(tableCell.CellFormat.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Ошибка в цвете заливки ячейки"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckTableCellFormatBorder(tableCell))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Ошибка в цвете заливки ячейки"
                );
                tableMistakes.Add(mistake);
            }

            if (!TableCellFormatTextDirection.Contains(tableCell.CellFormat.TextDirection))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Ошибка в направлении текста в ячейке"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckParagraphFormat(int tableRowIndex, int tableCellIndex, Word.Paragraph paragraph)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!BackgroundColor.Contains(paragraph.ParagraphFormat.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверный цвет заливки параграфа"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckParagraphFormatBorder(paragraph))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"У параграфа присутствуют рамки"
                );
                tableMistakes.Add(mistake);
            }

            if (!RightToLeft.Contains(paragraph.ParagraphFormat.RightToLeft))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Справа-налево'"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckParagraphCharacterFormat(int tableRowIndex, int tableCellIndex, Word.Paragraph paragraph)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (tableRowIndex == 0)
            {
                if (!FirstRowAlignment.Contains(paragraph.ParagraphFormat.Alignment))
                {
                    TableMistake mistake = new TableMistake(
                        row: tableRowIndex,
                        column: tableCellIndex,
                        message: $"Ошибка в выравнивании первой строки"
                    );
                    tableMistakes.Add(mistake);
                }
            }

            if ((tableCellIndex == 0) & (tableRowIndex != 0))
            {
                if (!FirstColumnAlignment.Contains(paragraph.ParagraphFormat.Alignment))
                {
                    TableMistake mistake = new TableMistake(
                        row: tableRowIndex,
                        column: tableCellIndex,
                        message: $"Ошибка в выравнивании первого столбца"
                    );
                    tableMistakes.Add(mistake);
                }
            }

            if (!WholeParagraphBackgroundColor.Contains(paragraph.CharacterFormatForParagraphMark.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет заливки' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckWholeParagraphBorder(paragraph))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"У параграфа присутствуют рамки"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphDoubleStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphFontColor.Contains(paragraph.CharacterFormatForParagraphMark.FontColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет шрифта' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphFontName.Contains(paragraph.CharacterFormatForParagraphMark.FontName))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Шрифт' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphHidden.Contains(paragraph.CharacterFormatForParagraphMark.Hidden))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Скрытый' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphHighlightColor.Contains(paragraph.CharacterFormatForParagraphMark.HighlightColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет выделения' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphKerning.Contains(paragraph.CharacterFormatForParagraphMark.Kerning))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Кернинг' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphRightToLeft.Contains(paragraph.CharacterFormatForParagraphMark.RightToLeft))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Справа-налево' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphScaling.Contains(paragraph.CharacterFormatForParagraphMark.Scaling))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Масштаб' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (paragraph.CharacterFormatForParagraphMark.Size != WholeParagraphChosenSize))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца (должно быть единообразие)"
                );
                tableMistakes.Add(mistake);
            }
            else
            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = paragraph.CharacterFormatForParagraphMark.Size;
            }

            if (!WholeParagraphStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.Strikethrough))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Зачеркнутый' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphSubscript.Contains(paragraph.CharacterFormatForParagraphMark.Subscript))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Подстрочный' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphSuperscript.Contains(paragraph.CharacterFormatForParagraphMark.Superscript))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Надстрочный' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            if (!WholeParagraphUnderlineStyle.Contains(paragraph.CharacterFormatForParagraphMark.UnderlineStyle))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Подчеркнутый' для всего абзаца"
                );
                tableMistakes.Add(mistake);
            }

            return tableMistakes;
        }

        private List<TableMistake> CheckRunnerCharacterFormat(int tableRowIndex, int tableCellIndex, Word.Run runner)
        {
            List<TableMistake> tableMistakes = new List<TableMistake>();

            if (!RunnerBackgroundColor.Contains(runner.CharacterFormat.BackgroundColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет заливки' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!CheckRunnerBorder(runner))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Границы' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerStrikethrough.Contains(runner.CharacterFormat.Strikethrough))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerFontColor.Contains(runner.CharacterFormat.FontColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет шрифта' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerFontName.Contains(runner.CharacterFormat.FontName))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Шрифт' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerHidden.Contains(runner.CharacterFormat.Hidden))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Скрытый' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerHighlightColor.Contains(runner.CharacterFormat.HighlightColor))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Цвет выделения' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerKerning.Contains(runner.CharacterFormat.Kerning))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Кернинг' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerRightToLeft.Contains(runner.CharacterFormat.RightToLeft))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Справа-налево' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerScaling.Contains(runner.CharacterFormat.Scaling))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Масштаб' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (runner.CharacterFormat.Size != WholeParagraphChosenSize))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера (должно быть единообразие)"
                );
                tableMistakes.Add(mistake);
            }
            else
            if ((runner.CharacterFormat.Size < RunnerSizeLeftBorder) | (runner.CharacterFormat.Size > RunnerSizeRightBorder))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера"
                );
                tableMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = runner.CharacterFormat.Size;
            }

            if (!RunnerStrikethrough.Contains(runner.CharacterFormat.Strikethrough))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Зачеркнутый' для раннера"
                );
                tableMistakes.Add(mistake);
            }

            if (!RunnerUnderlineStyle.Contains(runner.CharacterFormat.UnderlineStyle))
            {
                TableMistake mistake = new TableMistake(
                    row: tableRowIndex,
                    column: tableCellIndex,
                    message: $"Неверное значение свойства 'Подчеркнутый' для раннера"
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
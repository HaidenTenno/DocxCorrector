using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public abstract class DocumentElement
    {
        // Класс элемента
        public abstract ParagraphClass ParagraphClass { get; }

        // Свойства ParagraphFormat
        public virtual List<Word.HorizontalAlignment> Alignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Justify };
        public virtual List<Word.Color> BackgroundColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual List<Word.BorderStyle> BorderStyle => new List<Word.BorderStyle> { Word.BorderStyle.None };
        public virtual List<bool> KeepLinesTogether => new List<bool> { false };
        public virtual List<bool> KeepWithNext => new List<bool> { false };
        public virtual List<double> LeftIndentation => new List<double> { 0 };
        public virtual List<double> LineSpacing => new List<double> { 1.5 };
        public virtual List<Word.LineSpacingRule> LineSpacingRule => new List<Word.LineSpacingRule> { Word.LineSpacingRule.Multiple };
        public virtual List<bool> MirrorIndents => new List<bool> { false };
        public virtual List<bool> NoSpaceBetweenParagraphsOfSameStyle => new List<bool> { false };
        public virtual List<Word.OutlineLevel> OutlineLevel => new List<Word.OutlineLevel> { Word.OutlineLevel.BodyText };
        public virtual List<bool> PageBreakBefore => new List<bool> { false };
        public virtual List<double> RightIndentation => new List<double> { 0 };
        public virtual List<bool> RightToLeft => new List<bool> { false };
        public virtual List<double> SpaceAfter => new List<double> { 0 };
        public virtual List<double> SpaceBefore => new List<double> { 0 };
        public virtual double SpecialIndentationLeftBorder => -36.85;
        public virtual double SpecialIndentationRightBorder => -34.00;
        public virtual List<bool> WidowControl => new List<bool> { true };

        // Свойства CharacterFormat для всего абзаца
        public virtual List<bool> WholeParagraphAllCaps => new List<bool> { false };
        public virtual List<Word.Color> WholeParagraphBackgroundColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual List<bool> WholeParagraphBold => new List<bool> { false };
        public virtual List<Word.SingleBorder> WholeParagraphBorder => new List<Word.SingleBorder> { Word.SingleBorder.None };
        public virtual List<bool> WholeParagraphDoubleStrikethrough => new List<bool> { false };
        public virtual List<Word.Color> WholeParagraphFontColor => new List<Word.Color> { Word.Color.Black };
        public virtual List<string> WholeParagraphFontName => new List<string> { "Times New Roman" };
        public virtual List<bool> WholeParagraphHidden => new List<bool> { false };
        public virtual List<Word.Color> WholeParagraphHighlightColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual List<bool> WholeParagraphItalic => new List<bool> { false };
        public virtual List<double> WholeParagraphKerning => new List<double> { 0 };
        public virtual List<double> WholeParagraphPosition => new List<double> { 0 };
        public virtual List<bool> WholeParagraphRightToLeft => new List<bool> { false };
        public virtual List<int> WholeParagraphScaling => new List<int> { 100 };
        public virtual double WholeParagraphSizeLeftBorder => 13.5;
        public virtual double WholeParagraphSizeRightBorder => 14.5;
        public static double? WholeParagraphChosenSize { get; protected set; } = null;
        public virtual List<bool> WholeParagraphSmallCaps => new List<bool> { false };
        public virtual List<double> WholeParagraphSpacing => new List<double> { 0 };
        public virtual List<bool> WholeParagraphStrikethrough => new List<bool> { false };
        public virtual List<bool> WholeParagraphSubscript => new List<bool> { false };
        public virtual List<bool> WholeParagraphSuperscript => new List<bool> { false };
        public virtual List<Word.UnderlineType> WholeParagraphUnderlineStyle => new List<Word.UnderlineType> { Word.UnderlineType.None };

        // Свойства CharacterFormat для раннеров
        public virtual List<Word.Color> RunnerBackgroundColor => WholeParagraphBackgroundColor;
        public virtual List<Word.SingleBorder> RunnerBorder => WholeParagraphBorder;
        public virtual List<bool> RunnerDoubleStrikethrough => WholeParagraphDoubleStrikethrough;
        public virtual List<Word.Color> RunnerFontColor => WholeParagraphFontColor;
        public virtual List<string> RunnerFontName => WholeParagraphFontName;
        public virtual List<bool> RunnerHidden => WholeParagraphHidden;
        public virtual List<Word.Color> RunnerHighlightColor => WholeParagraphHighlightColor;
        public virtual List<double> RunnerKerning => WholeParagraphKerning;
        public virtual List<double> RunnerPosition => WholeParagraphPosition;
        public virtual List<bool> RunnerRightToLeft => WholeParagraphRightToLeft;
        public virtual List<int> RunnerScaling => WholeParagraphScaling;
        public virtual double RunnerSizeLeftBorder => WholeParagraphSizeLeftBorder;
        public virtual double RunnerSizeRightBorder => WholeParagraphSizeRightBorder;
        public virtual List<double> RunnerSpacing => WholeParagraphSpacing;
        public virtual List<bool> RunnerStrikethrough => WholeParagraphStrikethrough;
        public virtual List<Word.UnderlineType> RunnerUnderlineStyle => WholeParagraphUnderlineStyle;

        // Количество пустых строк (отбивок, SPACE, n0) после параграфа
        public virtual List<int> EmptyLinesAfter => new List<int> { 0 };

        // Проверка границ (Borders)
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
            if (!RunnerBorder.Contains(runner.CharacterFormat.Border))
            {
                return false;
            }
            return true;
        }

        private List<ParagraphMistake> CheckParagraphFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (!Alignment.Contains(paragraph.ParagraphFormat.Alignment))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное выравнивание"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!BackgroundColor.Contains(paragraph.ParagraphFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверный цвет заливки параграфа"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckParagraphFormatBorder(paragraph))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"У параграфа присутствуют рамки"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!KeepLinesTogether.Contains(paragraph.ParagraphFormat.KeepLinesTogether))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не разрывать абзац'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!KeepWithNext.Contains(paragraph.ParagraphFormat.KeepWithNext))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не отрывать от следующего'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!LeftIndentation.Contains(paragraph.ParagraphFormat.LeftIndentation))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение отступа слева"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!LineSpacing.Contains(paragraph.ParagraphFormat.LineSpacing))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение междустрочного интервала"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!LineSpacingRule.Contains(paragraph.ParagraphFormat.LineSpacingRule))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение типа междустрочного интервала"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!MirrorIndents.Contains(paragraph.ParagraphFormat.MirrorIndents))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зеркальные отступы'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!NoSpaceBetweenParagraphsOfSameStyle.Contains(paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не добавлять интервал между параграфами одного стиля'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!OutlineLevel.Contains(paragraph.ParagraphFormat.OutlineLevel))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение уровня заголовка"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!PageBreakBefore.Contains(paragraph.ParagraphFormat.PageBreakBefore))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'С новой страницы'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RightIndentation.Contains(paragraph.ParagraphFormat.RightIndentation))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Отступ справа'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RightToLeft.Contains(paragraph.ParagraphFormat.RightToLeft))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!SpaceAfter.Contains(paragraph.ParagraphFormat.SpaceAfter))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Интервал после'"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!SpaceBefore.Contains(paragraph.ParagraphFormat.SpaceBefore))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Интервал до'"
                );
                paragraphMistakes.Add(mistake);
            }

            // Отступ первой строки
            if ((paragraph.ParagraphFormat.SpecialIndentation < SpecialIndentationLeftBorder) | ((paragraph.ParagraphFormat.SpecialIndentation > SpecialIndentationRightBorder)))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение отступа первой строки"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WidowControl.Contains(paragraph.ParagraphFormat.WidowControl))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Запрет висячих строк'"
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private List<ParagraphMistake> CheckWholeParagraphCharacterFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (!WholeParagraphAllCaps.Contains(paragraph.CharacterFormatForParagraphMark.AllCaps))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все прописные' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphBackgroundColor.Contains(paragraph.CharacterFormatForParagraphMark.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphBold.Contains(paragraph.CharacterFormatForParagraphMark.Bold))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Жирный' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckWholeParagraphBorder(paragraph))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"У параграфа присутствует рамка"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphDoubleStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphFontColor.Contains(paragraph.CharacterFormatForParagraphMark.FontColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет шрифта' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphFontName.Contains(paragraph.CharacterFormatForParagraphMark.FontName))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Шрифт' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphHidden.Contains(paragraph.CharacterFormatForParagraphMark.Hidden))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphHighlightColor.Contains(paragraph.CharacterFormatForParagraphMark.HighlightColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphItalic.Contains(paragraph.CharacterFormatForParagraphMark.Italic))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Курсив' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphKerning.Contains(paragraph.CharacterFormatForParagraphMark.Kerning))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Кернинг' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphPosition.Contains(paragraph.CharacterFormatForParagraphMark.Position))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Смещение' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphRightToLeft.Contains(paragraph.CharacterFormatForParagraphMark.RightToLeft))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphScaling.Contains(paragraph.CharacterFormatForParagraphMark.Scaling))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Масштаб' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (paragraph.CharacterFormatForParagraphMark.Size != WholeParagraphChosenSize))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца (должно быть единообразие)"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = paragraph.CharacterFormatForParagraphMark.Size;
            }

            if (!WholeParagraphSmallCaps.Contains(paragraph.CharacterFormatForParagraphMark.SmallCaps))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все строчные' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphSpacing.Contains(paragraph.CharacterFormatForParagraphMark.Spacing))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Межсимвольный интервал' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.Strikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зачеркнутый' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphSubscript.Contains(paragraph.CharacterFormatForParagraphMark.Subscript))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подстрочный' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphSuperscript.Contains(paragraph.CharacterFormatForParagraphMark.Superscript))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Надстрочный' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphUnderlineStyle.Contains(paragraph.CharacterFormatForParagraphMark.UnderlineStyle))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подчеркнутый' для всего абзаца"
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private List<ParagraphMistake> CheckRunnerCharacterFormat(Word.Run runner)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (!RunnerBackgroundColor.Contains(runner.CharacterFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckRunnerBorder(runner))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Границы' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerDoubleStrikethrough.Contains(runner.CharacterFormat.DoubleStrikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerFontColor.Contains(runner.CharacterFormat.FontColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет шрифта' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerFontName.Contains(runner.CharacterFormat.FontName))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Шрифт' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerHidden.Contains(runner.CharacterFormat.Hidden))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerHighlightColor.Contains(runner.CharacterFormat.HighlightColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerKerning.Contains(runner.CharacterFormat.Kerning))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Кернинг' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerPosition.Contains(runner.CharacterFormat.Position))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Смещение' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerRightToLeft.Contains(runner.CharacterFormat.RightToLeft))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerScaling.Contains(runner.CharacterFormat.Scaling))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Масштаб' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (runner.CharacterFormat.Size != WholeParagraphChosenSize))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера (должно быть единообразие)"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if ((runner.CharacterFormat.Size < RunnerSizeLeftBorder) | (runner.CharacterFormat.Size > RunnerSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = runner.CharacterFormat.Size;
            }

            if (!RunnerSpacing.Contains(runner.CharacterFormat.Spacing))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Межсимвольный интервал' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerStrikethrough.Contains(runner.CharacterFormat.Strikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зачеркнутый' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerUnderlineStyle.Contains(runner.CharacterFormat.UnderlineStyle))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подчеркнутый' для раннера"
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private ParagraphMistake? CheckEmptyLines(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            // Посчитать количество строк до следующего параграфа
            int emptyLinesCount = 0;
            while (id + emptyLinesCount < classifiedParagraphs.Count)
            {
                emptyLinesCount++;
                int idToCheckEmpty = id + emptyLinesCount;

                Word.Paragraph paragraphToCheckForEmpty;
                // Если следующий элемент не параграф, то он не пустой
                try { paragraphToCheckForEmpty = (Word.Paragraph)classifiedParagraphs[idToCheckEmpty].Element; }
                catch { break; }

                if (GemBoxHelper.GetParagraphContentWithoutNewLine(paragraphToCheckForEmpty) != "") { break; }
            }

            if (!EmptyLinesAfter.Contains(emptyLinesCount - 1))
            {
                return new ParagraphMistake(
                    message: $"Неверное количество пропущенных параграфов"
                );
            }

            return null;
        }

        // Базовый метод проверки
        public virtual ParagraphCorrections? CheckFormatting(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Paragraph paragraph;
            // Если текущий элемент не параграф, то вернуть null
            try { paragraph = (Word.Paragraph)classifiedParagraphs[id].Element; } catch { return null; }

            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Свойства ParagraphFormat
            paragraphMistakes.AddRange(CheckParagraphFormat(paragraph));

            // Свойства CharacterFormat для всего абзаца
            paragraphMistakes.AddRange(CheckWholeParagraphCharacterFormat(paragraph));
            
            foreach (Word.Run runner in paragraph.GetChildElements(false, Word.ElementType.Run))
            {
                // Свойства CharacterFormat для раннеров
                paragraphMistakes.AddRange(CheckRunnerCharacterFormat(runner));
            }

            // Особые свойства
            // Проверка количества пустых строк
            ParagraphMistake? emptyLinesMistake = CheckEmptyLines(id, classifiedParagraphs);
            if (emptyLinesMistake != null) { paragraphMistakes.Add(emptyLinesMistake); }

            if (paragraphMistakes.Count != 0)
            {
                return new ParagraphCorrections(
                    paragraphID: id,
                    paragraphClass: ParagraphClass,
                    prefix: GemBoxHelper.GetParagraphPrefix(paragraph, 20),
                    mistakes: paragraphMistakes
                );
            }
            else
            {
                return null;
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public abstract class DocumentElement
    {
        // Класс элемента
        public abstract ParagraphClass ParagraphClass { get; }

        // Свойства ParagraphFormat
        public virtual List<Word.HorizontalAlignment> Alignment => new List<Word.HorizontalAlignment> { Word.HorizontalAlignment.Justify }; // 1)
        public virtual List<Word.Color> BackgroundColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White }; // 2)
        public virtual List<Word.BorderStyle> BorderStyle => new List<Word.BorderStyle> { Word.BorderStyle.None }; // 3)
        public virtual List<bool> KeepLinesTogether => new List<bool> { false }; // 4)
        public virtual List<bool> KeepWithNext => new List<bool> { false }; // 5)
        public virtual List<double> LeftIndentation => new List<double> { 0 }; // 6)
        public virtual List<double> LineSpacing => new List<double> { 1.5 }; // 7)
        public virtual List<Word.LineSpacingRule> LineSpacingRule => new List<Word.LineSpacingRule> { Word.LineSpacingRule.Multiple }; // 8)
        public virtual List<bool> MirrorIndents => new List<bool> { false }; // 9)
        public virtual List<bool> NoSpaceBetweenParagraphsOfSameStyle => new List<bool> { false }; // 10)
        public virtual List<Word.OutlineLevel> OutlineLevel => new List<Word.OutlineLevel> { Word.OutlineLevel.BodyText }; // 11)
        public virtual List<bool> PageBreakBefore => new List<bool> { false }; // 12)
        public virtual List<double> RightIndentation => new List<double> { 0 }; // 13)
        public virtual List<double> SpaceAfter => new List<double> { 0 }; // 14)
        public virtual List<double> SpaceBefore => new List<double> { 0 }; // 15)
        public virtual double SpecialIndentationLeftBorder => -36.85; // 16)
        public virtual double SpecialIndentationRightBorder => -34.00; // 16)
        public virtual List<bool> WidowControl => new List<bool> { true }; // 17)

        // Свойства CharacterFormat для всего абзаца
        public virtual List<bool> WholeParagraphAllCaps => new List<bool> { false }; // 18)
        public virtual List<Word.Color> WholeParagraphBackgroundColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White }; // 19)
        public virtual List<bool> WholeParagraphBold => new List<bool> { false }; // 20)
        public virtual List<Word.SingleBorder> WholeParagraphBorder => new List<Word.SingleBorder> { Word.SingleBorder.None }; // 21)
        public virtual List<bool> WholeParagraphDoubleStrikethrough => new List<bool> { false }; // 22)
        public virtual List<Word.Color> WholeParagraphFontColor => new List<Word.Color> { Word.Color.Black }; // 23)
        public virtual List<string> WholeParagraphFontName => new List<string> { "Times New Roman" }; // 24)
        public virtual List<bool> WholeParagraphHidden => new List<bool> { false }; // 25)
        public virtual List<Word.Color> WholeParagraphHighlightColor => new List<Word.Color> { Word.Color.Empty, Word.Color.White }; // 26)
        public virtual List<bool> WholeParagraphItalic => new List<bool> { false }; // 27)
        public virtual List<double> WholeParagraphKerning => new List<double> { 0 }; // 28)
        public virtual List<double> WholeParagraphPosition => new List<double> { 0 }; // 29)
        public virtual List<int> WholeParagraphScaling => new List<int> { 100 }; // 30)
        public virtual double WholeParagraphSizeLeftBorder => 13.5; // 31)
        public virtual double WholeParagraphSizeRightBorder => 14.5; // 31)
        public static double? WholeParagraphChosenSize { get; protected set; } = null; // 31
        public virtual List<bool> WholeParagraphSmallCaps => new List<bool> { false }; // 32)
        public virtual List<double> WholeParagraphSpacing => new List<double> { 0 }; // 33)
        public virtual List<bool> WholeParagraphStrikethrough => new List<bool> { false }; // 34)
        public virtual List<bool> WholeParagraphSubscript => new List<bool> { false }; // 35)
        public virtual List<bool> WholeParagraphSuperscript => new List<bool> { false }; // 36)
        public virtual List<Word.UnderlineType> WholeParagraphUnderlineStyle => new List<Word.UnderlineType> { Word.UnderlineType.None }; // 37)

        // Свойства CharacterFormat для раннеров
        public virtual List<Word.Color> RunnerBackgroundColor => WholeParagraphBackgroundColor; // 40
        public virtual List<Word.SingleBorder> RunnerBorder => WholeParagraphBorder; // 41
        public virtual List<bool> RunnerDoubleStrikethrough => WholeParagraphDoubleStrikethrough; // 42
        public virtual List<Word.Color> RunnerFontColor => WholeParagraphFontColor; // 43
        public virtual List<string> RunnerFontName => WholeParagraphFontName; // 44
        public virtual List<bool> RunnerHidden => WholeParagraphHidden; // 45
        public virtual List<Word.Color> RunnerHighlightColor => WholeParagraphHighlightColor; // 46
        public virtual List<double> RunnerKerning => WholeParagraphKerning; // 47
        public virtual List<double> RunnerPosition => WholeParagraphPosition; // 48
        public virtual List<int> RunnerScaling => WholeParagraphScaling; // 50
        public virtual double RunnerSizeLeftBorder => WholeParagraphSizeLeftBorder; // 51
        public virtual double RunnerSizeRightBorder => WholeParagraphSizeRightBorder; // 51
        public virtual List<double> RunnerSpacing => WholeParagraphSpacing; // 52
        public virtual List<bool> RunnerStrikethrough => WholeParagraphStrikethrough; // 53
        public virtual List<Word.UnderlineType> RunnerUnderlineStyle => WholeParagraphUnderlineStyle; // 54

        // Количество пустых строк (отбивок, SPACE, n0) после параграфа
        public virtual List<int> EmptyLinesAfter => new List<int> { 0 };

        // Проверка, что список не пустой
        private bool CheckIfListIsNotEmpty<T>(List<T> list)
        {
            return list.Count() != 0;
        }

        // Проверка границ (Borders)
        private bool CheckParagraphFormatBorder(Word.Paragraph paragraph)
        {
            foreach (Word.SingleBorderType borderType in Enum.GetValues(typeof(Word.SingleBorderType)))
            {
                if (CheckIfListIsNotEmpty(BorderStyle) & !BorderStyle.Contains(paragraph.ParagraphFormat.Borders[borderType].Style))
                {
                    return false;
                }
            }
            return true;
        }

        private bool CheckWholeParagraphBorder(Word.Paragraph paragraph)
        {
            if (CheckIfListIsNotEmpty(WholeParagraphBorder) & !WholeParagraphBorder.Contains(paragraph.CharacterFormatForParagraphMark.Border))
            {
                return false;
            }
            return true;
        }

        private bool CheckRunnerBorder(Word.Run runner)
        {
            if (CheckIfListIsNotEmpty(RunnerBorder) & !RunnerBorder.Contains(runner.CharacterFormat.Border))
            {
                return false;
            }
            return true;
        }

        private List<ParagraphMistake> CheckParagraphFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (CheckIfListIsNotEmpty(Alignment) & !Alignment.Contains(paragraph.ParagraphFormat.Alignment))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное выравнивание",
                    advice: AdviceCreator.ParagraphAligment(Alignment)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(BackgroundColor) & !BackgroundColor.Contains(paragraph.ParagraphFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверный цвет заливки параграфа",
                    advice: AdviceCreator.BackgroundColor(BackgroundColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckParagraphFormatBorder(paragraph))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"У параграфа присутствуют рамки",
                    advice: AdviceCreator.BorderStyle(BorderStyle)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(KeepLinesTogether) & !KeepLinesTogether.Contains(paragraph.ParagraphFormat.KeepLinesTogether))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не разрывать абзац'",
                    advice: AdviceCreator.KeepLineTogether(KeepLinesTogether)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(KeepWithNext) & !KeepWithNext.Contains(paragraph.ParagraphFormat.KeepWithNext))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не отрывать от следующего'",
                    advice: AdviceCreator.KeepWithNext(KeepWithNext)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(LeftIndentation) & !LeftIndentation.Contains(paragraph.ParagraphFormat.LeftIndentation))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение отступа слева",
                    advice: AdviceCreator.LeftIdentation(LeftIndentation)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(LineSpacing) & !LineSpacing.Contains(paragraph.ParagraphFormat.LineSpacing))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение междустрочного интервала",
                    advice: AdviceCreator.LineSpacing(LineSpacing)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(MirrorIndents) & !MirrorIndents.Contains(paragraph.ParagraphFormat.MirrorIndents))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зеркальные отступы'",
                    advice: AdviceCreator.MirrorIndents(MirrorIndents)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(NoSpaceBetweenParagraphsOfSameStyle) & !NoSpaceBetweenParagraphsOfSameStyle.Contains(paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не добавлять интервал между параграфами одного стиля'",
                    advice: AdviceCreator.NoSpaceBetweenParagraphsOfSameStyle(NoSpaceBetweenParagraphsOfSameStyle)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(OutlineLevel) & !OutlineLevel.Contains(paragraph.ParagraphFormat.OutlineLevel))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение уровня заголовка",
                    advice: AdviceCreator.OutlineLevel(OutlineLevel)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(PageBreakBefore) & !PageBreakBefore.Contains(paragraph.ParagraphFormat.PageBreakBefore))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'С новой страницы'",
                    advice: AdviceCreator.PageBreakBefore(PageBreakBefore)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RightIndentation) & !RightIndentation.Contains(paragraph.ParagraphFormat.RightIndentation))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Отступ справа'",
                    advice: AdviceCreator.RightIndentation(RightIndentation)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(SpaceAfter) & !SpaceAfter.Contains(paragraph.ParagraphFormat.SpaceAfter))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Интервал после'",
                    advice: AdviceCreator.SpaceAfter(SpaceAfter)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(SpaceBefore) & !SpaceBefore.Contains(paragraph.ParagraphFormat.SpaceBefore))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Интервал до'",
                    advice: AdviceCreator.SpaceAfter(SpaceBefore)
                );
                paragraphMistakes.Add(mistake);
            }

            // Отступ первой строки
            if ((paragraph.ParagraphFormat.SpecialIndentation < SpecialIndentationLeftBorder) | ((paragraph.ParagraphFormat.SpecialIndentation > SpecialIndentationRightBorder)))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение отступа первой строки",
                    advice: AdviceCreator.SpecialIndentation(SpecialIndentationLeftBorder, SpecialIndentationRightBorder)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WidowControl) & !WidowControl.Contains(paragraph.ParagraphFormat.WidowControl))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Запрет висячих строк'",
                    advice: AdviceCreator.WidowControl(WidowControl)
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private List<ParagraphMistake> CheckWholeParagraphCharacterFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (CheckIfListIsNotEmpty(WholeParagraphAllCaps) & !WholeParagraphAllCaps.Contains(paragraph.CharacterFormatForParagraphMark.AllCaps))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все прописные' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphAllCaps(WholeParagraphAllCaps)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphBackgroundColor) & !WholeParagraphBackgroundColor.Contains(paragraph.CharacterFormatForParagraphMark.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphBackgroundColor(WholeParagraphBackgroundColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphBold) & !WholeParagraphBold.Contains(paragraph.CharacterFormatForParagraphMark.Bold))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Жирный' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphBold(WholeParagraphBold)
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckWholeParagraphBorder(paragraph))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"У параграфа присутствует рамка",
                    advice: AdviceCreator.WholeParagraphBorder(WholeParagraphBorder)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphDoubleStrikethrough) & !WholeParagraphDoubleStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphDoubleStrikethrough(WholeParagraphDoubleStrikethrough)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphFontColor) & !WholeParagraphFontColor.Contains(paragraph.CharacterFormatForParagraphMark.FontColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет шрифта' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphFontColor(WholeParagraphFontColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphFontName) & !WholeParagraphFontName.Contains(paragraph.CharacterFormatForParagraphMark.FontName))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Шрифт' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphFontName(WholeParagraphFontName)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphHidden) & !WholeParagraphHidden.Contains(paragraph.CharacterFormatForParagraphMark.Hidden))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphHidden(WholeParagraphHidden)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphHighlightColor) & !WholeParagraphHighlightColor.Contains(paragraph.CharacterFormatForParagraphMark.HighlightColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphHighlightColor(WholeParagraphHighlightColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphItalic) & !WholeParagraphItalic.Contains(paragraph.CharacterFormatForParagraphMark.Italic))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Курсив' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphItalic(WholeParagraphItalic)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphKerning) & !WholeParagraphKerning.Contains(paragraph.CharacterFormatForParagraphMark.Kerning))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Кернинг' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphKerning(WholeParagraphKerning)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphPosition) & !WholeParagraphPosition.Contains(paragraph.CharacterFormatForParagraphMark.Position))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Смещение' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphPosition(WholeParagraphPosition)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphScaling) & !WholeParagraphScaling.Contains(paragraph.CharacterFormatForParagraphMark.Scaling))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Масштаб' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphScaling(WholeParagraphScaling)
                );
                paragraphMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (paragraph.CharacterFormatForParagraphMark.Size != WholeParagraphChosenSize))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца (должно быть единообразие)",
                    advice: AdviceCreator.WholeParagraphChosenSize((double)WholeParagraphChosenSize!)
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphSizeBorder(WholeParagraphSizeLeftBorder, WholeParagraphSizeRightBorder)
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = paragraph.CharacterFormatForParagraphMark.Size;
            }

            if (CheckIfListIsNotEmpty(WholeParagraphSmallCaps) & !WholeParagraphSmallCaps.Contains(paragraph.CharacterFormatForParagraphMark.SmallCaps))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все строчные' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphSmallCaps(WholeParagraphSmallCaps)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphSpacing) & !WholeParagraphSpacing.Contains(paragraph.CharacterFormatForParagraphMark.Spacing))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Межсимвольный интервал' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphSpacing(WholeParagraphSpacing)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphStrikethrough) & !WholeParagraphStrikethrough.Contains(paragraph.CharacterFormatForParagraphMark.Strikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зачеркнутый' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphStrikethrough(WholeParagraphStrikethrough)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphSubscript) & !WholeParagraphSubscript.Contains(paragraph.CharacterFormatForParagraphMark.Subscript))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подстрочный' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphSubscript(WholeParagraphSubscript)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphSuperscript) & !WholeParagraphSuperscript.Contains(paragraph.CharacterFormatForParagraphMark.Superscript))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Надстрочный' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphSuperscript(WholeParagraphSuperscript)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(WholeParagraphUnderlineStyle) & !WholeParagraphUnderlineStyle.Contains(paragraph.CharacterFormatForParagraphMark.UnderlineStyle))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подчеркнутый' для всего абзаца",
                    advice: AdviceCreator.WholeParagraphUnderlineStyle(WholeParagraphUnderlineStyle)
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private List<ParagraphMistake> CheckRunnerCharacterFormat(Word.Run runner)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (CheckIfListIsNotEmpty(RunnerBackgroundColor) & !RunnerBackgroundColor.Contains(runner.CharacterFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для раннера",
                    advice: AdviceCreator.WholeParagraphBackgroundColor(RunnerBackgroundColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckRunnerBorder(runner))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Границы' для раннера",
                    advice: AdviceCreator.WholeParagraphBorder(RunnerBorder)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerDoubleStrikethrough) & !RunnerDoubleStrikethrough.Contains(runner.CharacterFormat.DoubleStrikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для раннера",
                    advice: AdviceCreator.WholeParagraphDoubleStrikethrough(RunnerDoubleStrikethrough)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerFontColor) & !RunnerFontColor.Contains(runner.CharacterFormat.FontColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет шрифта' для раннера",
                    advice: AdviceCreator.WholeParagraphFontColor(RunnerFontColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerFontName) & !RunnerFontName.Contains(runner.CharacterFormat.FontName))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Шрифт' для раннера",
                    advice: AdviceCreator.WholeParagraphFontName(RunnerFontName)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerHidden) & !RunnerHidden.Contains(runner.CharacterFormat.Hidden))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для раннера",
                    advice: AdviceCreator.WholeParagraphHidden(RunnerHidden)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerHighlightColor) & !RunnerHighlightColor.Contains(runner.CharacterFormat.HighlightColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для раннера",
                    advice: AdviceCreator.WholeParagraphHighlightColor(RunnerHighlightColor)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerKerning) & !RunnerKerning.Contains(runner.CharacterFormat.Kerning))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Кернинг' для раннера",
                    advice: AdviceCreator.WholeParagraphKerning(RunnerKerning)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerPosition) & !RunnerPosition.Contains(runner.CharacterFormat.Position))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Смещение' для раннера",
                    advice: AdviceCreator.WholeParagraphPosition(RunnerPosition)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerScaling) & !RunnerScaling.Contains(runner.CharacterFormat.Scaling))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Масштаб' для раннера",
                    advice: AdviceCreator.WholeParagraphScaling(RunnerScaling)
                );
                paragraphMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (runner.CharacterFormat.Size != WholeParagraphChosenSize))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера (должно быть единообразие)",
                    AdviceCreator.WholeParagraphChosenSize((double)WholeParagraphChosenSize!)
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if ((runner.CharacterFormat.Size < RunnerSizeLeftBorder) | (runner.CharacterFormat.Size > RunnerSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера",
                    advice: AdviceCreator.WholeParagraphSizeBorder(RunnerSizeLeftBorder, RunnerSizeRightBorder)
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = runner.CharacterFormat.Size;
            }

            if (CheckIfListIsNotEmpty(RunnerSpacing) & !RunnerSpacing.Contains(runner.CharacterFormat.Spacing))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Межсимвольный интервал' для раннера",
                    advice: AdviceCreator.WholeParagraphSpacing(RunnerSpacing)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerStrikethrough) & !RunnerStrikethrough.Contains(runner.CharacterFormat.Strikethrough))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зачеркнутый' для раннера",
                    advice: AdviceCreator.WholeParagraphDoubleStrikethrough(RunnerStrikethrough)
                );
                paragraphMistakes.Add(mistake);
            }

            if (CheckIfListIsNotEmpty(RunnerUnderlineStyle) & !RunnerUnderlineStyle.Contains(runner.CharacterFormat.UnderlineStyle))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подчеркнутый' для раннера",
                    advice: AdviceCreator.WholeParagraphUnderlineStyle(RunnerUnderlineStyle)
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

        // Проверить все раннеры параграфа
        private List<ParagraphMistake> CheckRunners(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            var runners = paragraph.GetChildElements(false, Word.ElementType.Run).ToList();
            //if (runners.Count == 1) { return paragraphMistakes; }

            foreach (Word.Run runner in runners)
            {
                // Свойства CharacterFormat для раннеров
                paragraphMistakes.AddRange(CheckRunnerCharacterFormat(runner));
            }

            return paragraphMistakes;
        }

        // Проверить особые свойства
        protected virtual List<ParagraphMistake> CheckSpecialFeatures(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            ParagraphMistake? emptyLinesMistake = CheckEmptyLines(id, classifiedParagraphs);
            if (emptyLinesMistake != null) { paragraphMistakes.Add(emptyLinesMistake); }

            return paragraphMistakes;
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

            // Свойства CharacterFormat для раннеров
            paragraphMistakes.AddRange(CheckRunners(paragraph));

            // Особые свойства
            // Проверка количества пустых строк
            paragraphMistakes.AddRange(CheckSpecialFeatures(id, classifiedParagraphs));
            
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

        // Выполнить сравнение по свойствам (не включая особые) с параграфом paragraph
        public virtual ParagraphCorrections? CheckSingleParagraphFormatting(int id, Word.Paragraph paragraph)
        {
            // TODO: Подумать над проверкой выбранного шрифта
            ResetChosenSize();

            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Свойства ParagraphFormat
            paragraphMistakes.AddRange(CheckParagraphFormat(paragraph));

            // Свойства CharacterFormat для всего абзаца
            paragraphMistakes.AddRange(CheckWholeParagraphCharacterFormat(paragraph));

            // Свойства CharacterFormat для раннеров
            paragraphMistakes.AddRange(CheckRunners(paragraph));

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

        // TODO: Убрать?
        public void ResetChosenSize()
        {
            WholeParagraphChosenSize = null;
        } 
    }
}
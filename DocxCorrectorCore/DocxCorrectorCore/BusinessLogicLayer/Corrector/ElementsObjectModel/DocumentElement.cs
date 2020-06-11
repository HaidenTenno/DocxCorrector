using System;
using System.Collections.Generic;
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
        public virtual Word.HorizontalAlignment Alignment => Word.HorizontalAlignment.Justify;
        public virtual List<Word.Color> BackgroundColors => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual Word.BorderStyle BorderStyle => Word.BorderStyle.None;
        public virtual bool KeepLinesTogether => false;
        public virtual bool KeepWithNext => false;
        public virtual double LeftIndentation => 0;
        public virtual double LineSpacing => 1.5;
        public virtual Word.LineSpacingRule LineSpacingRule => Word.LineSpacingRule.Multiple;
        public virtual bool MirrorIndents => false;
        public virtual bool NoSpaceBetweenParagraphsOfSameStyle => false;
        public virtual Word.OutlineLevel OutlineLevel => Word.OutlineLevel.BodyText;
        public virtual bool PageBreakBefore => false;
        public virtual double RightIndentation => 0;
        public virtual bool RightToLeft => false;
        public virtual double SpaceAfter => 0;
        public virtual double SpaceBefore => 0;
        public virtual double SpecialIndentationLeftBorder => -36.85;
        public virtual double SpecialIndentationRightBorder => -35.85;
        public virtual bool WidowControl => true;

        // Свойства CharacterFormat для всего абзаца
        public virtual bool WholeParagraphAllCaps => false;
        public virtual List<Word.Color> WholeParagraphBackgroundColors => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual bool WholeParagraphBold => false;
        public virtual Word.SingleBorder WholeParagraphBorder => Word.SingleBorder.None;
        public virtual bool WholeParagraphDoubleStrikethrough => false;
        public virtual Word.Color WholeParagraphFontColor => Word.Color.Black;
        public virtual string WholeParagraphFontName => "Times New Roman";
        public virtual bool WholeParagraphHidden => false;
        public virtual List<Word.Color> WholeParagraphHighlightColors => new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual bool WholeParagraphItalic => false;
        public virtual double WholeParagraphKerning => 0;
        public virtual double WholeParagraphPosition => 0;
        public virtual bool WholeParagraphRightToLeft => false;
        public virtual int WholeParagraphScaling => 100;
        public virtual double WholeParagraphSizeLeftBorder => 13.5;
        public virtual double WholeParagraphSizeRightBorder => 14.5;
        public static double? WholeParagraphChosenSize { get; protected set; } = null;
        public virtual bool WholeParagraphSmallCaps => false;
        public virtual double WholeParagraphSpacing => 0;
        public virtual bool WholeParagraphStrikethrough => false;
        public virtual bool WholeParagraphSubscript => false;
        public virtual bool WholeParagraphSuperscript => false;
        public virtual Word.UnderlineType WholeParagraphUnderlineStyle => Word.UnderlineType.None;

        // Свойства CharacterFormat для раннеров
        public virtual List<Word.Color> RunnerBackgroundColors => WholeParagraphBackgroundColors;//new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual Word.SingleBorder RunnerBorder => WholeParagraphBorder;//Word.SingleBorder.None;
        public virtual bool RunnerDoubleStrikethrough => WholeParagraphDoubleStrikethrough;//false;
        public virtual Word.Color RunnerFontColor => WholeParagraphFontColor;//Word.Color.Black;
        public virtual string RunnerFontName => WholeParagraphFontName;//"Times New Roman";
        public virtual bool RunnerHidden => WholeParagraphHidden;//false;
        public virtual List<Word.Color> RunnerHighlightColors => WholeParagraphHighlightColors;//new List<Word.Color> { Word.Color.Empty, Word.Color.White };
        public virtual double RunnerKerning => WholeParagraphKerning;//0;
        public virtual double RunnerPosition => WholeParagraphPosition;//0;
        public virtual bool RunnerRightToLeft => WholeParagraphRightToLeft;//false;
        public virtual int RunnerScaling => WholeParagraphScaling;//100;
        public virtual double RunnerSizeLeftBorder => WholeParagraphSizeLeftBorder;//13.5;
        public virtual double RunnerSizeRightBorder => WholeParagraphSizeRightBorder;//14.5;
        public virtual double RunnerSpacing => WholeParagraphSpacing;//0;
        public virtual bool RunnerStrikethrough => WholeParagraphStrikethrough;//false;
        public virtual Word.UnderlineType RunnerUnderlineStyle => WholeParagraphUnderlineStyle;//Word.UnderlineType.None;

        // Количество пустых строк (отбивок, SPACE, n0) после параграфа
        public virtual int EmptyLinesAfter => 0;

        // Проверка границ (Borders)
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

        private List<ParagraphMistake> CheckParagraphFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (paragraph.ParagraphFormat.Alignment != Alignment)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное выравнивание",
                    advice: $"Выбрано {paragraph.ParagraphFormat.Alignment}; Требуется {Alignment}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!BackgroundColors.Contains(paragraph.ParagraphFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверный цвет заливки параграфа",
                    advice: $"Выбрано {paragraph.ParagraphFormat.BackgroundColor}; Требуется {BackgroundColors}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckParagraphFormatBorder(paragraph))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"У параграфа присутствуют рамки",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.KeepLinesTogether != KeepLinesTogether)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не разрывать абзац'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.KeepLinesTogether}; Требуется {KeepLinesTogether}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.KeepWithNext != KeepWithNext)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не отрывать от следующего'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.KeepWithNext}; Требуется {KeepWithNext}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.LeftIndentation != LeftIndentation)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение отступа слева",
                    advice: $"Выбрано {paragraph.ParagraphFormat.LeftIndentation}; Требуется {LeftIndentation}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.LineSpacing != LineSpacing)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение междустрочного интервала",
                    advice: $"Выбрано {paragraph.ParagraphFormat.LineSpacing}; Требуется {LineSpacing}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.LineSpacingRule != LineSpacingRule)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение типа междустрочного интервала",
                    advice: $"Выбрано {paragraph.ParagraphFormat.LineSpacingRule}; Требуется {LineSpacingRule}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.MirrorIndents != MirrorIndents)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зеркальные отступы'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.MirrorIndents}; Требуется {MirrorIndents}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle != NoSpaceBetweenParagraphsOfSameStyle)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Не добавлять интервал между параграфами одного стиля'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle}; Требуется {NoSpaceBetweenParagraphsOfSameStyle}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.OutlineLevel != OutlineLevel)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение уровня заголовка",
                    advice: $"Выбрано {paragraph.ParagraphFormat.OutlineLevel}; Требуется {OutlineLevel}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.PageBreakBefore != PageBreakBefore)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'С новой страницы'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.PageBreakBefore}; Требуется {PageBreakBefore}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.RightIndentation != RightIndentation)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Отступ справа'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.RightIndentation}; Требуется {RightIndentation}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.RightToLeft != RightToLeft)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.RightToLeft}; Требуется {RightToLeft}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.SpaceAfter != SpaceAfter)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Интервал после'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.SpaceAfter}; Требуется {SpaceAfter}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.SpaceBefore != SpaceBefore)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Интервал до'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.SpaceBefore}; Требуется {SpaceBefore}"
                );
                paragraphMistakes.Add(mistake);
            }

            // Отступ первой строки
            if ((paragraph.ParagraphFormat.SpecialIndentation < SpecialIndentationLeftBorder) | ((paragraph.ParagraphFormat.SpecialIndentation > SpecialIndentationRightBorder)))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение отступа первой строки",
                    advice: $"Выбрано {paragraph.ParagraphFormat.SpecialIndentation}; Требуется значение между {SpecialIndentationLeftBorder} и {SpecialIndentationRightBorder}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.WidowControl != WidowControl)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Запрет висячих строк'",
                    advice: $"Выбрано {paragraph.ParagraphFormat.WidowControl}; Требуется {WidowControl}"
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private List<ParagraphMistake> CheckWholeParagraphCharacterFormat(Word.Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (paragraph.CharacterFormatForParagraphMark.AllCaps != WholeParagraphAllCaps)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все прописные' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.AllCaps}; Требуется {WholeParagraphAllCaps}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphBackgroundColors.Contains(paragraph.CharacterFormatForParagraphMark.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.BackgroundColor}; Требуется {WholeParagraphBackgroundColors}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Bold != WholeParagraphBold)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Жирный' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Bold}; Требуется {WholeParagraphBold}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckWholeParagraphBorder(paragraph))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"У параграфа присутствует рамка",
                    advice: $"Тут будет совет"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough != WholeParagraphDoubleStrikethrough)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.DoubleStrikethrough}; Требуется {WholeParagraphDoubleStrikethrough}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.FontColor != WholeParagraphFontColor)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет шрифта' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.FontColor}; Требуется {WholeParagraphFontColor}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.FontName != WholeParagraphFontName)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Шрифт' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.FontName}; Требуется {WholeParagraphFontName}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Hidden != WholeParagraphHidden)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Hidden}; Требуется {WholeParagraphHidden}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!WholeParagraphHighlightColors.Contains(paragraph.CharacterFormatForParagraphMark.HighlightColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.HighlightColor}; Требуется {WholeParagraphHighlightColors}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Italic != WholeParagraphItalic)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Курсив' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Italic}; Требуется {WholeParagraphItalic}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Kerning != WholeParagraphKerning)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Кернинг' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Kerning}; Требуется {WholeParagraphKerning}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Position != WholeParagraphPosition)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Смещение' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Position}; Требуется {WholeParagraphPosition}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.RightToLeft != WholeParagraphRightToLeft)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.RightToLeft}; Требуется {WholeParagraphRightToLeft}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Scaling != WholeParagraphScaling)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Масштаб' для всего абзаца",
                    advice: "ТУТ БУДЕТ СОВЕТ"
                );
                paragraphMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (paragraph.CharacterFormatForParagraphMark.Size != WholeParagraphChosenSize))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца (должно быть единообразие)",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Size}; Требуется значение между {WholeParagraphSizeLeftBorder} и {WholeParagraphSizeRightBorder}"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = paragraph.CharacterFormatForParagraphMark.Size;
            }

            if (paragraph.CharacterFormatForParagraphMark.SmallCaps != WholeParagraphSmallCaps)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все строчные' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.SmallCaps}; Требуется {WholeParagraphSmallCaps}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Spacing != WholeParagraphSpacing)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Межсимвольный интервал' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Spacing}; Требуется {WholeParagraphSpacing}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Strikethrough != WholeParagraphStrikethrough)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зачеркнутый' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Strikethrough}; Требуется {WholeParagraphStrikethrough}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Subscript != WholeParagraphSubscript)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подстрочный' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Subscript}; Требуется {WholeParagraphSubscript}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.Superscript != WholeParagraphSuperscript)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Надстрочный' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Superscript}; Требуется {WholeParagraphSuperscript}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.UnderlineStyle != WholeParagraphUnderlineStyle)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подчеркнутый' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.UnderlineStyle}; Требуется {WholeParagraphUnderlineStyle}"
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private List<ParagraphMistake> CheckRunnerCharacterFormat(Word.Run runner)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            if (!RunnerBackgroundColors.Contains(runner.CharacterFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.BackgroundColor}; Требуется {RunnerBackgroundColors}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!CheckRunnerBorder(runner))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Границы' для раннера",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.DoubleStrikethrough != RunnerDoubleStrikethrough)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Двойное зачеркивание' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.DoubleStrikethrough}; Требуется {RunnerDoubleStrikethrough}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.FontColor != RunnerFontColor)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет шрифта' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.FontColor}; Требуется {RunnerFontColor}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.FontName != RunnerFontName)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Шрифт' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.FontName}; Требуется {RunnerFontName}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Hidden != RunnerHidden)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Hidden}; Требуется {RunnerHidden}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (!RunnerHighlightColors.Contains(runner.CharacterFormat.HighlightColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.HighlightColor}; Требуется {RunnerHighlightColors}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Kerning != RunnerKerning)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Кернинг' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Kerning}; Требуется {RunnerKerning}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Position != RunnerPosition)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Смещение' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Position}; Требуется {RunnerPosition}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.RightToLeft != RunnerRightToLeft)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.RightToLeft}; Требуется {RunnerRightToLeft}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Scaling != RunnerScaling)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Масштаб' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Scaling}; Требуется {RunnerScaling}"
                );
                paragraphMistakes.Add(mistake);
            }

            // Проверка размера шрифта
            if ((WholeParagraphChosenSize != null) & (runner.CharacterFormat.Size != WholeParagraphChosenSize))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера (должно быть единообразие)",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if ((runner.CharacterFormat.Size < RunnerSizeLeftBorder) | (runner.CharacterFormat.Size > RunnerSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Size}; Требуется значение между {RunnerSizeLeftBorder} и {RunnerSizeRightBorder}"
                );
                paragraphMistakes.Add(mistake);
            }
            else
            if (WholeParagraphChosenSize == null)
            {
                WholeParagraphChosenSize = runner.CharacterFormat.Size;
            }

            if (runner.CharacterFormat.Spacing != RunnerSpacing)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Межсимвольный интервал' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Spacing}; Требуется {RunnerSpacing}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.Strikethrough != RunnerStrikethrough)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Зачеркнутый' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.Strikethrough}; Требуется {RunnerStrikethrough}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (runner.CharacterFormat.UnderlineStyle != RunnerUnderlineStyle)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Подчеркнутый' для раннера",
                    advice: $"Выбрано {runner.CharacterFormat.UnderlineStyle}; Требуется {RunnerUnderlineStyle}"
                );
                paragraphMistakes.Add(mistake);
            }

            return paragraphMistakes;
        }

        private ParagraphMistake? CheckEmptyLines(int id, List<ClassifiedParagraph> classifiedParagraphs)
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
                    return new ParagraphMistake(
                        message: $"Неверное количество пропущенных параграфов (недостаточно)",
                        advice: $"ТУТ БУДЕТ СОВЕТ"
                    );
                }

                if (GemBoxHelper.GetParagraphContentWithoutNewLine(paragraphToCheckForEmpty) != "")
                {
                    return new ParagraphMistake(
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
                return new ParagraphMistake(
                    message: $"Неверное количество пропущенных параграфов (слишком много)",
                    advice: $"ТУТ БУДЕТ СОВЕТ"
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
using DocxCorrectorCore.Services.Helpers;
using GemBox.Document;
using System.Collections.Generic;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public enum StartSymbolType : int
    {
        Upper,
        Lower,
        Number,
        Other
    }

    public abstract class DocumentElement
    {
        // Класс элемента
        public abstract ParagraphClass ParagraphClass { get; }

        // Свойства ParagraphFormat
        public abstract HorizontalAlignment Alignment { get; }
        public Color BackgroundColor => Color.Empty;
        public MultipleBorders? Borders => null; // TODO: Разобраться с получением свойств границ параграфа
        public bool KeepLinesTogether => false;
        public abstract bool KeepWithNext { get; }
        public double LeftIndentation => 0;
        public double LineSpacing => 1.5;
        public LineSpacingRule LineSpacingRule => LineSpacingRule.Multiple;
        public bool MirrorIndents => false;
        public bool NoSpaceBetweenParagraphsOfSameStyle => false;
        public abstract OutlineLevel OutlineLevel { get; }
        public abstract bool PageBreakBefore { get; }
        public double RightIndentation => 0;
        public bool RightToLeft => false;
        public double SpaceAfter => 0;
        public double SpaceBefore => 0;
        public abstract double SpecialIndentationLeftBorder { get; }
        public abstract double SpecialIndentationRightBorder { get; }
        public virtual ParagraphStyle? Style => null;
        public bool WidowControl => true;
        
        // Свойства CharacterFormat для всего абзаца
        public abstract bool WholeParagraphAllCaps { get; }
        public Color WholeParagraphBackgroundColor => Color.Empty;
        public abstract bool WholeParagraphBold { get; }
        public SingleBorder WholeParagraphBorder => SingleBorder.None;
        public bool WholeParagraphDoubleStrikethrough => false;
        public Color WholeParagraphFontColor => Color.Black;
        public string WholeParagraphFontName => "TimesNewRoman";
        public bool WholeParagraphHidden => false;
        public Color WholeParagraphHighlightColor => Color.Empty;
        public bool WholeParagraphItalic => false;
        public double WholeParagraphKerning => 0;
        public double WholeParagraphPosition => 0;
        public bool WholeParagraphRightToLeft => false;
        public int WholeParagraphScaling => 100; // TODO: Проверить, что это проценты
        public double WholeParagraphSizeLeftBorder => 13.5;
        public double WholeParagraphSizeRightBorder => 14.5;
        public abstract bool WholeParagraphSmallCaps { get; }
        public double WholeParagraphSpacing => 0;
        public bool WholeParagraphStrikethrough => false;
        public virtual CharacterStyle? WholeParagraphStyle => null;
        public bool WholeParagraphSubscript => false;
        public bool WholeParagraphSuperscript => false;
        public Color? WholeParagraphUnderlineColor => null;
        public UnderlineType WholeParagraphUnderlineStyle => UnderlineType.None;

        // Свойства CharacterFormat для всего абзаца
        public bool? RunnerAllCaps => null;
        public Color RunnerBackgroundColor => Color.Empty;
        public abstract bool RunnerBold { get; }
        public SingleBorder RunnerBorder => SingleBorder.None;
        public bool RunnerDoubleStrikethrough => false;
        public Color RunnerFontColor => Color.Black;
        public string RunnerFontName => "TimesNewRoman";
        public bool RunnerHidden => false;
        public Color RunnerHighlightColor => Color.Empty;
        public bool? RunnerItalic => null;
        public double RunnerKerning => 0;
        public double RunnerPosition => 0;
        public bool RunnerRightToLeft => false;
        public int RunnerScaling => 100; // TODO: Проверить, что это проценты
        public double RunnerSizeLeftBorder => 13.5;
        public double RunnerSizeRightBorder => 14.5;
        public bool? RunnerSmallCaps => null;
        public double RunnerSpacing => 0;
        public bool RunnerStrikethrough => false;
        public virtual CharacterStyle? RunnerStyle => null;
        public bool? RunnerSubscript => null;
        public bool? RunnerSuperscript => null;
        public Color? RunnerUnderlineColor => null;
        public UnderlineType RunnerUnderlineStyle => UnderlineType.None;
        
        // Особые свойства
        // Особенность начального символа
        public virtual StartSymbolType? StartSymbol => null;
        
        // Префиксы
        public virtual string[]? Prefixes => null;
        
        // Суффиксы
        public virtual string[]? Suffixes => null;
        
        // Количество пустых строк (отбивок, SPACE, n0) после параграфа
        public abstract int EmptyLinesAfter { get; }

        // Базовый метод проверки
        public virtual ParagraphCorrections? CheckFormatting(int id, Paragraph paragraph)
        {
            List<ParagraphMistake> paragraphMistakes = new List<ParagraphMistake>();

            // Свойства ParagraphFormat
            if (paragraph.ParagraphFormat.Alignment != Alignment)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное выравнивание",
                    advice: $"Выбрано {paragraph.ParagraphFormat.Alignment}; Требуется {Alignment}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.ParagraphFormat.BackgroundColor != BackgroundColor)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверный цвет заливки параграфа",
                    advice: $"Выбрано {paragraph.ParagraphFormat.BackgroundColor}; Требуется {BackgroundColor}"
                );
                paragraphMistakes.Add(mistake);
            }

            // TODO: Border

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

            // IMPORTANT ОТСТУП ПЕРВОЙ СТРОКИ
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

            // Свойства CharacterFormat для всего абзаца
            if (paragraph.CharacterFormatForParagraphMark.AllCaps != WholeParagraphAllCaps)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Все прописные' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.AllCaps}; Требуется {WholeParagraphAllCaps}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.BackgroundColor != WholeParagraphBackgroundColor)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет заливки' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.BackgroundColor}; Требуется {WholeParagraphBackgroundColor}"
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

            if (paragraph.CharacterFormatForParagraphMark.Hidden != WholeParagraphHidden)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Скрытый' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Hidden}; Требуется {WholeParagraphHidden}"
                );
                paragraphMistakes.Add(mistake);
            }

            if (paragraph.CharacterFormatForParagraphMark.HighlightColor != WholeParagraphHighlightColor)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Цвет выделения' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.HighlightColor}; Требуется {WholeParagraphHighlightColor}"
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

            if (paragraph.CharacterFormatForParagraphMark.RightToLeft != WholeParagraphRightToLeft)
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Справа-налево' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.RightToLeft}; Требуется {WholeParagraphRightToLeft}"
                );
                paragraphMistakes.Add(mistake);
            }

            // TODO: ДОДЕЛАТЬ НАЧИНАЯ СО SCALING

            // Свойства CharacterFormat для всего абзаца

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
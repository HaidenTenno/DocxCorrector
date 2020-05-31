using System.Collections.Generic;
using GemBox.Document;
using DocxCorrectorCore.Services.Helpers;

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
        public List<Color> BackgroundColors => new List<Color> { Color.Empty, Color.White };

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
        public bool WidowControl => true;
        
        // Свойства CharacterFormat для всего абзаца
        public abstract bool WholeParagraphAllCaps { get; }
        public List<Color> WholeParagraphBackgroundColors => new List<Color> { Color.Empty, Color.White };
        public abstract bool WholeParagraphBold { get; }
        public SingleBorder WholeParagraphBorder => SingleBorder.None;
        public bool WholeParagraphDoubleStrikethrough => false;
        public Color WholeParagraphFontColor => Color.Black;
        public string WholeParagraphFontName => "Times New Roman";
        public bool WholeParagraphHidden => false;
        public List<Color> WholeParagraphHighlightColors => new List<Color> { Color.Empty, Color.White };
        public bool WholeParagraphItalic => false;
        public double WholeParagraphKerning => 0;
        public double WholeParagraphPosition => 0;
        public bool WholeParagraphRightToLeft => false;
        public int WholeParagraphScaling => 100;
        public double WholeParagraphSizeLeftBorder => 13.5;
        public double WholeParagraphSizeRightBorder => 14.5;
        public abstract bool WholeParagraphSmallCaps { get; }
        public double WholeParagraphSpacing => 0;
        public bool WholeParagraphStrikethrough => false;
        public bool WholeParagraphSubscript => false;
        public bool WholeParagraphSuperscript => false;
        public Color? WholeParagraphUnderlineColor => null;
        public UnderlineType WholeParagraphUnderlineStyle => UnderlineType.None;
        public bool? RunnerAllCaps => null;
        public List<Color> RunnerBackgroundColors => new List<Color> { Color.Empty, Color.White };
        public abstract bool RunnerBold { get; }
        public SingleBorder RunnerBorder => SingleBorder.None;
        public bool RunnerDoubleStrikethrough => false;
        public Color RunnerFontColor => Color.Black;
        public string RunnerFontName => "Times New Roman";
        public bool RunnerHidden => false;
        public List<Color> RunnerHighlightColors => new List<Color> { Color.Empty, Color.White };
        public bool? RunnerItalic => null;
        public double RunnerKerning => 0;
        public double RunnerPosition => 0;
        public bool RunnerRightToLeft => false;
        public int RunnerScaling => 100;
        public double RunnerSizeLeftBorder => 13.5;
        public double RunnerSizeRightBorder => 14.5;
        public bool? RunnerSmallCaps => null;
        public double RunnerSpacing => 0;
        public bool RunnerStrikethrough => false;
        public bool? RunnerSubscript => null;
        public bool? RunnerSuperscript => null;
        public UnderlineType RunnerUnderlineStyle => UnderlineType.None;

        // TODO: Возможно нужно убать большинсмтво особых свойств из Document Element

        //// Особые свойства
        ////Особенность начального символа
        //public virtual StartSymbolType? StartSymbol => null;

        //// Префиксы
        //public virtual string[]? Prefixes => null;
        
        //// Суффиксы
        //public virtual string[]? Suffixes => null;
        
        // Количество пустых строк (отбивок, SPACE, n0) после параграфа
        public abstract int EmptyLinesAfter { get; }

        // Базовый метод проверки
        public virtual ParagraphCorrections? CheckFormatting(int id, List<Paragraph> paragraphs)
        {
            Paragraph paragraph;
            try { paragraph = paragraphs[id]; } catch { return null; }

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

            if (!BackgroundColors.Contains(paragraph.ParagraphFormat.BackgroundColor))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверный цвет заливки параграфа",
                    advice: $"Выбрано {paragraph.ParagraphFormat.BackgroundColor}; Требуется {BackgroundColors}"
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

            if ((paragraph.CharacterFormatForParagraphMark.Size < WholeParagraphSizeLeftBorder) | (paragraph.CharacterFormatForParagraphMark.Size > WholeParagraphSizeRightBorder))
            {
                ParagraphMistake mistake = new ParagraphMistake(
                    message: $"Неверное значение свойства 'Размер шрифта' для всего абзаца",
                    advice: $"Выбрано {paragraph.CharacterFormatForParagraphMark.Size}; Требуется значение между {WholeParagraphSizeLeftBorder} и {WholeParagraphSizeRightBorder}"
                );
                paragraphMistakes.Add(mistake);
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

            // Свойства CharacterFormat для раннеров
            foreach (Run runner in paragraph.GetChildElements(false, ElementType.Run))
            {
                // AllCaps?
                
                if (!RunnerBackgroundColors.Contains(runner.CharacterFormat.BackgroundColor))
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверное значение свойства 'Цвет заливки' для раннера",
                        advice: $"Выбрано {runner.CharacterFormat.BackgroundColor}; Требуется {RunnerBackgroundColors}"
                    );
                    paragraphMistakes.Add(mistake);
                }
                
                if (runner.CharacterFormat.Bold != RunnerBold)
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверное значение свойства 'Жирный' для раннера",
                        advice: $"Выбрано {runner.CharacterFormat.Bold}; Требуется {RunnerBold}"
                    );
                    paragraphMistakes.Add(mistake);
                }
                
                if (runner.CharacterFormat.Border != RunnerBorder)
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверное значение свойства 'Границы' для раннера",
                        advice: $"Выбрано {runner.CharacterFormat.Border}; Требуется {RunnerBorder}"
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
                
                // Italic?
                
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
                
                if ((runner.CharacterFormat.Size < RunnerSizeLeftBorder) | (runner.CharacterFormat.Size > RunnerSizeRightBorder))
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверное значение свойства 'Размер шрифта' для раннера",
                        advice: $"Выбрано {runner.CharacterFormat.Size}; Требуется значение между {RunnerSizeLeftBorder} и {RunnerSizeRightBorder}"
                    );
                    paragraphMistakes.Add(mistake);
                }
                
                // SmallCaps?
                
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
                
                // Subscript?
                
                // Superscript?
                
                if (runner.CharacterFormat.UnderlineStyle != RunnerUnderlineStyle)
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверное значение свойства 'Подчеркнутый' для раннера",
                        advice: $"Выбрано {runner.CharacterFormat.UnderlineStyle}; Требуется {RunnerUnderlineStyle}"
                    );
                    paragraphMistakes.Add(mistake);
                }
            }

            // Особые свойства
            // Количество пустых строк (отбивок, SPACE, n0) после параграфа
            int emptyLinesCount = 1;
            while ((emptyLinesCount <= EmptyLinesAfter) & (id + emptyLinesCount < paragraphs.Count))
            {
                int idToCheckEmpty = id + emptyLinesCount;
                Paragraph paragraphToCheckForEmpty = paragraphs[idToCheckEmpty];

                string paragraphToCheckEmptyContent = "";
                foreach (Run runner in paragraphToCheckForEmpty.GetChildElements(false, ElementType.Run))
                {
                    paragraphToCheckEmptyContent += runner.Content;
                }
                paragraphToCheckEmptyContent = paragraphToCheckEmptyContent.Trim();

                if (paragraphToCheckEmptyContent != "")
                {
                    ParagraphMistake mistake = new ParagraphMistake(
                        message: $"Неверное количество пропущенных параграфов",
                        advice: $"ТУТ БУДЕТ СОВЕТ"
                    );
                    paragraphMistakes.Add(mistake);
                    break;
                }
                emptyLinesCount++;
            }

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
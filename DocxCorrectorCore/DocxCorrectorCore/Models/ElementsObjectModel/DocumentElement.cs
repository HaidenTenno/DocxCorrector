using GemBox.Document;

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
    }
}

namespace DocxCorrector.Models
{
    // Выравнивание
    public enum NormalizedAligment: int
    {
        Left,
        Center,
        Right,
        Justify,
        Other
    }

    // TODO: - Naming
    // Для проверки на наличия слов, оформленных жирным, курсивом и т. д.
    public enum ContainsStatus: int
    {
        None, // Параграф не содержит таких слов
        Contains, // Параграф содержит слова с таким оформлением
        Full // Параграф оформлен полностью таким спосом
    }

    // Правило междустрочного интервала
    public enum LineSpacingRule: int
    {
        Single,
        OneAndHalf,
        Double,
        Miltiply,
        Other
    }


    public class NormalizedProperties
    {
        // ID параграфа
        public int Id { get; set; }
        // Отступ красной строки
        public float FirstLineIndent { get; set; }
        // Выравнивание
        public int Aligment { get; set; }
        // Количество символов
        public int SymbolsCount { get; set; }
        // Префикс - число
        public int PrefixIsNumber { get; set; }
        // Префикс - строчная буква 
        public int PrefixIsLowercase { get; set; }
        // Префикс - прописная буква
        public int PrefixIsUppercase { get; set; }
        // Префикс - ["-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－"]
        public int PrefixIsDash { get; set; }
        // Суффикс - [".","!","?"]
        public int SuffixIsEndSign { get; set; }
        // Суффикс - двоеточие
        public int SuffixIsSemicolon { get; set; }
        // Суффикс - запятая или точка с запятой
        public int SuffixIsCommaOrSemicolon { get; set; }
        // Содержит - ["-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－"]
        public int ContainsDash { get; set; }
        // Содержит ")"
        public int ContainsBracket { get; set; }
        // Размер шрифта
        public float FontSize { get; set; }
        // Междустрочный интервал
        public float LineSpacing { get; set; }
        // Правило междустрочного интервала
        public int LineSpacingRule { get; set; }
        // Курсив
        public int Italic { get; set; }
        // Жирный
        public int Bold { get; set; }
        // Цвет черный
        public int BlackColor { get; set; }
    }
}

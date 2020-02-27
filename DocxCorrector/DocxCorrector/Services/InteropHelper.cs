using System;
using Word = Microsoft.Office.Interop.Word;

namespace DocxCorrector.Services
{
    internal static class InteropHelper
    {
        // Проверить, что первый символ абзаца принадлежит множеству символов
        internal static int CheckIfFirstSymbolOfParagraphIs(Word.Paragraph paragraph, string[] symbols)
        {
            return Array.IndexOf(symbols, paragraph.Range.Text[0].ToString()) != -1 ? 1 : 0;
        }

        // Проверить, что последний символ абзаца принадлежит можнеству символов
        internal static int CheckIfLastSymbolOfParagraphIs(Word.Paragraph paragraph, string[] symbols)
        {
            if (paragraph.Range.Text.Length > 1)
            {
                return Array.IndexOf(symbols, paragraph.Range.Text[paragraph.Range.Text.Length - 2].ToString()) != -1 ? 1 : 0;
            }
            else
            {
                return CheckIfFirstSymbolOfParagraphIs(paragraph, symbols);
            }
        }

        // Проверить, что параграф содержит хотя бы один из символов
        internal static int CheckIfParagraphsContainsOneOf(Word.Paragraph paragraph, string[] symbols)
        {
            foreach (string symbol in symbols)
            {
                if (paragraph.Range.Text.Contains(symbol))
                {
                    return 1;
                }
            }
            return 0;
        }
    }
}
using System;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Services.Helpers
{
    internal static class GemBoxHelper
    {
        // Ввод лицензионного ключа
        internal static void SetLicense()
        {
            Word.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        // Открыть документ filePath
        internal static Word.DocumentModel? OpenDocument(string filePath)
        {
            try
            {
                Word.DocumentModel document = Word.DocumentModel.Load(filePath);
                document.CalculateListItems();
                document.GetPaginator(new Word.PaginatorOptions() { UpdateFields = true });
                return document;
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                Console.WriteLine("Can't open document");
                return null;
            }
        }

        // Проверить, что первый символ абзаца принадлежит множеству символов
        internal static int CheckIfFirstSymbolOfParagraphIs(Word.Paragraph paragraph, string[] symbols)
        {
            return Array.IndexOf(symbols, paragraph.Content.ToString()[0].ToString()) != -1 ? 1 : 0;
        }

        // Проверить, что последний символ абзаца принадлежит можнеству символов
        internal static int CheckIfLastSymbolOfParagraphIs(Word.Paragraph paragraph, string[] symbols)
        {
            if (paragraph.Content.ToString().Length > 2)
            {
                return Array.IndexOf(symbols, paragraph.Content.ToString()[paragraph.Content.ToString().Length - 3].ToString()) != -1 ? 1 : 0;
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
                if (paragraph.Content.ToString().Contains(symbol))
                {
                    return 1;
                }
            }
            return 0;
        }

        // Получить первые prefixLength символов параграфа paragraph (если длина меньшье, то вернуть весь параграф)
        internal static string GetParagraphPrefix(Word.Paragraph paragraph, int prefixLength)
        {
            return paragraph.Content.ToString().Length > prefixLength ? paragraph.Content.ToString().Substring(0, prefixLength) : paragraph.Content.ToString();
        }
    }
}

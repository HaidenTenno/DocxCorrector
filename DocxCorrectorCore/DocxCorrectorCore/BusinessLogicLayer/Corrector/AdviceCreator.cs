using System;
using System.Collections.Generic;
using System.Text;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector
{
    internal static class AdviceCreator
    {
        // Private
        private static bool CheckListEmpty<T> (List<T> list)
        {
            return list.Count == 0;
        }

        // Public
        // 1
        public static string ParagraphAligment(List<Word.HorizontalAlignment> alignments)
        {
            if (CheckListEmpty(alignments)) return "Значение не определено";
            string prefix = alignments.Count > 1 ? "Выравнивание текста должно быть одним из следующих: " : "Выравнивание текста должно быть ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.HorizontalAlignment alignment in alignments)
            {
                switch (alignment)
                {
                    case Word.HorizontalAlignment.Center:
                        possibleStrings.Add("по центру");
                        break;
                    case Word.HorizontalAlignment.Justify:
                        possibleStrings.Add("по ширине");
                        break;
                    case Word.HorizontalAlignment.Left:
                        possibleStrings.Add("по левому краю");
                        break;
                    case Word.HorizontalAlignment.Right:
                        possibleStrings.Add("по правому краю");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 2
        public static string BackgroundColor(List<Word.Color> colors)
        {
            if (CheckListEmpty(colors)) return "Значение не определено";
            string prefix = colors.Count > 1 ? "Цвет заливки параграфа должнен быть одним из следующих: " : "Цвет заливки параграфа должнен быть ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.Color color in colors)
            {
                switch (color)
                {
                    case { } clr when clr == Word.Color.White:
                        possibleStrings.Add("белый");
                        break;
                    case { } clr when clr == Word.Color.Empty:
                        possibleStrings.Add("без заливки ('нет цвета')");
                        break;
                    default:
                        possibleStrings.Add("неизвестный цвет");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 3
        public static string BorderStyle(List<Word.BorderStyle> styles)
        {
            if (CheckListEmpty(styles)) return "Значение не определено";
            string prefix = styles.Count > 1 ? "Стиль рамок параграфа должнен быть одним из следующих: " : "Стиль рамок параграфа должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.BorderStyle style in styles)
            {
                switch (style)
                {
                    case Word.BorderStyle.None:
                        possibleStrings.Add("баз рамок");
                        break;
                    default:
                        possibleStrings.Add("неизвестный стиль");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 4
        public static string KeepLineTogether(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = (values[0] == true) ? "Свойство 'не разрывать абзац' должно быть включено" :
                "Свойство 'не разрывать абзац' должно быть выключено";
            return advice;
        }

        // 5
        public static string KeepWithNext(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'не отрывать от следующего' должно быть включено" : "Свойство 'не отрывать от следующего' должно быть выключено";
            return advice;
        }

        // 6
        public static string LeftIdentation(List<double> indentations)
        {
            if (CheckListEmpty(indentations)) return "Значение не определено";
            string prefix = indentations.Count > 1 ? "Отступ слева должнен быть одним из следующих: " : "Отступ слева должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double indentation in indentations)
            {
                possibleStrings.Add($"{indentation}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 7
        public static string LineSpacing(List<double> spaces)
        {
            if (CheckListEmpty(spaces)) return "Значение не определено";
            string prefix = spaces.Count > 1 ? "Междустрочный интервал должнен быть одним из следующих: " : "Междустрочный интервал должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double space in spaces)
            {
                possibleStrings.Add($"{space}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }
        
        // 8
        public static string LineSpacingRule(List<Word.LineSpacingRule> rules)
        {
            if (CheckListEmpty(rules)) return "Значение не определено";
            string prefix = rules.Count > 1 ? "Тип междустрочного интервала должнен быть одним из следующих: " : "Тип междустрочного интервала должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.LineSpacingRule rule in rules)
            {
                switch (rule)
                {
                    case Word.LineSpacingRule.Multiple:
                        possibleStrings.Add("множитель");
                        break;
                    case Word.LineSpacingRule.AtLeast:
                        possibleStrings.Add("точно");
                        break;
                    default:
                        possibleStrings.Add("неизвестно");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }
    }
}

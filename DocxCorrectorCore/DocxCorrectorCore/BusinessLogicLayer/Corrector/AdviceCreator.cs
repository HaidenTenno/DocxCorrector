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
                        possibleStrings.Add("требуется уточнение интервала");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 9
        public static string MirrorIndents(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Зеркальные отступы' должно быть включено" : "Свойство 'Зеркальные отступы' должно быть выключено";
            return advice;
        }

        // 10
        public static string NoSpaceBetweenParagraphsOfSameStyle(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Не добавлять интервал между параграфами одного стиля' должно быть включено" : "Свойство 'Не добавлять интервал между параграфами одного стиля' должно быть выключено";
            return advice;
        }

        // 11
        public static string OutlineLevel(List<Word.OutlineLevel> levels)
        {
            if (CheckListEmpty(levels)) return "Значение не определено";
            string prefix = levels.Count > 1 ? "Уровень заголовка должнен быть одним из следующих: " : "Уровень заголовка должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.OutlineLevel level in levels)
            {
                switch (level)
                {
                    case Word.OutlineLevel.BodyText:
                        possibleStrings.Add("основной текст");
                        break;
                    case Word.OutlineLevel.Level1:
                        possibleStrings.Add("уровень 1");
                        break;
                    case Word.OutlineLevel.Level2:
                        possibleStrings.Add("уровень 2");
                        break;
                    case Word.OutlineLevel.Level3:
                        possibleStrings.Add("уровень 3");
                        break;
                    default:
                        possibleStrings.Add("требуется уточнение уровня");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 12
        public static string PageBreakBefore(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'С новой страницы' должно быть включено" : "Свойство 'С новой страницы' должно быть выключено";
            return advice;
        }

        // 13
        public static string RightIndentation(List<double> indentations)
        {
            if (CheckListEmpty(indentations)) return "Значение не определено";
            string prefix = indentations.Count > 1 ? "Отступ справа должнен быть одним из следующих: " : "Отступ справа должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double indentation in indentations)
            {
                possibleStrings.Add($"{indentation}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 14
        public static string RightToLeft(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Слева-направо' должно быть включено" : "Свойство 'Слева-направо' должно быть выключено";
            return advice;
        }

        // 15
        public static string SpaceAfter(List<double> spaces)
        {
            if (CheckListEmpty(spaces)) return "Значение не определено";
            string prefix = spaces.Count > 1 ? "Интервал после должнен быть одним из следующих: " : "Интервал после должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double space in spaces)
            {
                possibleStrings.Add($"{space}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 16
        public static string SpaceBefore(List<double> spaces)
        {
            if (CheckListEmpty(spaces)) return "Значение не определено";
            string prefix = spaces.Count > 1 ? "Интервал до должнен быть одним из следующих: " : "Интервал до должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double space in spaces)
            {
                possibleStrings.Add($"{space}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 17
        public static string SpecialIndentation(double leftBorder, double rightBorder)
        {
            double wordValue1 = Math.Round((Math.Abs(rightBorder) * 1.25 / 35.85), 2, MidpointRounding.AwayFromZero);
            double wordValue2 = Math.Round((Math.Abs(leftBorder) * 1.25 / 35.85), 2, MidpointRounding.AwayFromZero);

            string advice = $"Значение отступа первой строки должно быть в пределах: {wordValue1} см - {wordValue2} см";
            return advice;
        }

        // 18
        public static string WidowControl(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Запрет висячих строк' должно быть включено" : "Свойство 'Запрет висячих строк' должно быть выключено";
            return advice;
        }

        // 19
        public static string WholeParagraphAllCaps(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Все прописные' должно быть включено" : "Свойство 'Все прописные' должно быть выключено";
            return advice;
        }

        // 20
        public static string WholeParagraphBackgroundColor(List<Word.Color> colors)
        {
            if (CheckListEmpty(colors)) return "Значение не определено";
            string prefix = colors.Count > 1 ? "Цвет заливки фона должнен быть одним из следующих: " : "Цвет заливки документа должнен быть ";

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

        // 21
        public static string WholeParagraphBold(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Жирный' должно быть включено" : "Свойство 'Жирный' должно быть выключено";
            return advice;
        }

        // 22
        public static string WholeParagraphBorder(List<Word.SingleBorder> styles)
        {
            if (CheckListEmpty(styles)) return "Значение не определено";
            string prefix = styles.Count > 1 ? "Стиль рамок должнен быть одним из следующих: " : "Стиль рамок документа должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.SingleBorder style in styles)
            {

                switch (style)
                {
                    case { } brdr when brdr == Word.SingleBorder.None:
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

        // 23
        public static string WholeParagraphDoubleStrikethrough(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Двойное зачеркивание' должно быть включено" : "Свойство 'Двойное зачеркивание' должно быть выключено";
            return advice;
        }

        // 24
        public static string WholeParagraphFontColor(List<Word.Color> colors)
        {
            if (CheckListEmpty(colors)) return "Значение не определено";
            string prefix = colors.Count > 1 ? "Цвет шрифта параграфа должнен быть одним из следующих: " : "Цвет шрифта параграфа должнен быть ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.Color color in colors)
            {
                switch (color)
                {
                    case { } clr when clr == Word.Color.Black:
                        possibleStrings.Add("чёрный");
                        break;
                    default:
                        possibleStrings.Add("неизвестный цвет");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 25
        public static string WholeParagraphFontName(List<string> fonts)
        {
            if (CheckListEmpty(fonts)) return "Значение не определено";
            string prefix = fonts.Count > 1 ? "Имя шрифта должнено быть одним из следующих: " : "Имя шрифта документа должнено быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (string font in fonts)
            {
                switch (font)
                {
                    case "Times New Roman":
                        possibleStrings.Add("Times New Roman");
                        break;
                    case "Courier New":
                        possibleStrings.Add("Courier New");
                        break;
                    case "Arial":
                        possibleStrings.Add("Arial");
                        break;
                    case "Calibri":
                        possibleStrings.Add("Calibri");
                        break;
                    default:
                        possibleStrings.Add("неизвестный шрифт");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }
        

        // 26
        public static string WholeParagraphHidden(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Скрытый' должно быть включено" : "Свойство 'Скрытый' должно быть выключено";
            return advice;
        }

        // 27
        public static string WholeParagraphHighlightColor(List<Word.Color> colors)
        {
            if (CheckListEmpty(colors)) return "Значение не определено";
            string prefix = colors.Count > 1 ? "Цвет выделения параграфа должнен быть одним из следующих: " : "Цвет выделения параграфа должнен быть ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.Color color in colors)
            {
                switch (color)
                {
                    case { } clr when clr == Word.Color.White:
                        possibleStrings.Add("белый");
                        break;
                    default:
                        possibleStrings.Add("неизвестный цвет");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 28
        public static string WholeParagraphItalic(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Курсив' должно быть включено" : "Свойство 'Курсив' должно быть выключено";
            return advice;
        }

        // 29
        public static string WholeParagraphKerning(List<double> kernings)
        {
            if (CheckListEmpty(kernings)) return "Значение не определено";
            string prefix = kernings.Count > 1 ? "Кернинг должнен быть одним из следующих: " : "Кернинг должнен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double kerning in kernings)
            {
                possibleStrings.Add($"{kerning}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 30
        public static string WholeParagraphPosition(List<double> positions)
        {
            if (CheckListEmpty(positions)) return "Значение не определено";
            string prefix = positions.Count > 1 ? "Смещение должно быть одним из следующих: " : "Смещение должно быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double position in positions)
            {
                possibleStrings.Add($"{position}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 31
        public static string WholeParagraphRightToLeft(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Слева-направо' должно быть включено" : "Свойство 'Слева-направо' должно быть выключено";
            return advice;
        }

        // 32
        public static string WholeParagraphScaling(List<int> scalings)
        {
            string advice;
            advice = scalings[0] == 100 ? "Свойство 'Масштаб 100' должно быть включено" : "Свойство 'Масштаб 100' должно быть выключено";
            return advice;
        }

        // 33
        public static string WholeParagraphSizeBorder(double leftBorder, double rightBorder)
        {
            string advice = $"Значение размера шрифта должно быть в пределах: {leftBorder} - {rightBorder}";
            return advice;
        }

        public static string WholeParagraphChosenSize(double chosenSize)
        {
            string advice = $"Значение размера шрифта должно быть {chosenSize}";
            return advice;
        }

        // 34
        public static string WholeParagraphSmallCaps(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Малые прописные' должно быть включено" : "Свойство 'Малые прописные' должно быть выключено";
            return advice;
        }

        // 35
        public static string WholeParagraphSpacing(List<double> spacings)
        {
            if (CheckListEmpty(spacings)) return "Значение не определено";
            string prefix = spacings.Count > 1 ? "Интервал между буквами должен быть одним из следующих: " : "Интервал между буквами должен быть: ";

            List<string> possibleStrings = new List<string>();
            foreach (double spacing in spacings)
            {
                possibleStrings.Add($"{spacing}");
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

        // 36
        public static string WholeParagraphStrikethrough(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Зачеркивание' должно быть включено" : "Свойство 'Зачеркивание' должно быть выключено";
            return advice;
        }

        // 37
        public static string WholeParagraphSubscript(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Подстрочный' должно быть включено" : "Свойство 'Подстрочный' должно быть выключено";
            return advice;
        }

        // 38
        public static string WholeParagraphSuperscript(List<bool> values)
        {
            if (CheckListEmpty(values)) return "Значение не определено";
            string advice;
            advice = values[0] == true ? "Свойство 'Надстрочный' должно быть включено" : "Свойство 'Надстрочный' должно быть выключено";
            return advice;
        }

        // 39
        public static string WholeParagraphUnderlineStyle(List<Word.UnderlineType> UnderlineTypes)
        {
            if (CheckListEmpty(UnderlineTypes)) return "Значение не определено";
            string prefix = UnderlineTypes.Count > 1 ? "Стиль подчеркивания должнен быть одним из следующих: " : "Цвет выделения документа должнен быть ";

            List<string> possibleStrings = new List<string>();
            foreach (Word.UnderlineType UnderlineType in UnderlineTypes)
            {
                switch (UnderlineType)
                {
                    case Word.UnderlineType.None:
                        possibleStrings.Add("нет подчёркивания");
                        break;
                    default:
                        possibleStrings.Add("неизвестный стиль");
                        break;
                }
            }

            string advice = prefix + string.Join(", ", possibleStrings);
            return advice;
        }

    }
}

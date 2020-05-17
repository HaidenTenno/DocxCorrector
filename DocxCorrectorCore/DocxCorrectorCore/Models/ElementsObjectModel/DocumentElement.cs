﻿using System;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public enum AligmentType : int
    {
        Left,
        Center,
        Right,
        Justify,
        Other
    }

    public enum StartSymbolType : int
    {
        Upper,
        Lower,
        Number,
        Other
    }

    public abstract class DocumentElement
    {
        // Отступ сверху
        public virtual int SpaceBefore => 0;

        // Отступ снизу
        public virtual int SpaceAfter => 0;

        // Множитель междустрочного интервала
        public virtual float LineSpacingMultiplier => 1.5f;

        // Название шрифта
        public virtual string FontName => "TimesNewRoman";

        // Размер шрифта
        public virtual float FontSize => 14f;

        // Отступ первой строки
        public virtual float FirstLineIndent => 35.45f;

        // Курсив
        public virtual bool Italic => false;

        // Жирность
        public virtual bool Bold => false;

        // Подчеркнутость
        public virtual bool Underline => false;

        // Зачеркнутость
        public virtual bool StrikeThrough => false;

        // Выделенность цветом
        public virtual int? HighLighted => null;

        // Цвет текста
        public virtual int TextColor => 0xFFFFFF;

        // Выравнивание
        public virtual AligmentType Aligment => AligmentType.Justify;

        // Особенность начального символа
        public virtual StartSymbolType StartSymbol => StartSymbolType.Upper;

        // Префиксы
        public virtual string[]? Prefixes => null;

        // Суффиксы
        public virtual string[]? Suffixes => null;

        // Не отрывать от следующего
        public virtual bool KeepWithNext => false;
    }
}
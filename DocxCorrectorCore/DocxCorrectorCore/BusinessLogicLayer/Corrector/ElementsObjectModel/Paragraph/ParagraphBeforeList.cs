﻿using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class ParagraphBeforeList : DocumentElement
    {
        //c2

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c2;

        // Свойства ParagraphFormat
        public override bool KeepLinesTogether => true;

        // Свойства CharacterFormat для всего абзаца

        // Свойства CharacterFormat для всего абзаца
        
        // Особые свойства
    }
}

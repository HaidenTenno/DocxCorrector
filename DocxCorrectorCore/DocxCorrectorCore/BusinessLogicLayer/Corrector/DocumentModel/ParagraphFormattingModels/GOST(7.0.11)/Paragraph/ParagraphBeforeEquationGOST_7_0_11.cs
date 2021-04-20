using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ParagraphBeforeEquationGOST_7_0_11 : DocumentElementGOST_7_0_11
    {
        //c3

        // Класс элемента
        public override ParagraphClass ParagraphClass => ParagraphClass.c3;

        // Свойства ParagraphFormat
        public override List<bool> KeepLinesTogether => new List<bool> { true };
        
        // Свойства CharacterFormat для всего абзаца
        
        // Свойства CharacterFormat для всего абзаца
        
        // Особые свойства
    }
}
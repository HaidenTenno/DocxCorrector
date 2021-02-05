using System;
using System.Collections.Generic;
using System.Text;
using DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel.GOST_7_32
{
    public class ParagraphFormattingModelGOST_7_32 : ParagraphFormattingModel
    {
        public override DocumentElement? GetDocumentElementFromClass(ParagraphClass paragraphClass)
        {
            return paragraphClass switch
            {
                ParagraphClass.b1 => new HeadingFirstLevel(),
                ParagraphClass.b2 => new HeadingOtherLevels(ParagraphClass.b2),
                ParagraphClass.b3 => new HeadingOtherLevels(ParagraphClass.b3),
                ParagraphClass.b4 => new HeadingOtherLevels(ParagraphClass.b4),
                ParagraphClass.c1 => new ParagraphRegular(),
                ParagraphClass.c2 => new ParagraphBeforeList(),
                ParagraphClass.c3 => new ParagraphBeforeEquation(),
                ParagraphClass.d1 => new SimpleListFirstElement(),
                ParagraphClass.d2 => new SimpleListMiddleElement(),
                ParagraphClass.d3 => new SimpleListLastElement(),
                ParagraphClass.d4 => new ComplexListFirstElement(),
                ParagraphClass.d5 => new ComplexListMiddleElement(),
                ParagraphClass.d6 => new ComplexListLastElement(),
                ParagraphClass.e0 => new Table(),
                ParagraphClass.f1 => new TableSign(ParagraphClass.f1),
                ParagraphClass.f3 => new TableSign(ParagraphClass.f3),
                ParagraphClass.h1 => new ImageSign(),
                _ => null
            };
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.DocumentModel
{
    public class ParagraphFormattingModelGOST_7_0_11 : ParagraphFormattingModel
    {
        public override DocumentElement? GetDocumentElementFromClass(ParagraphClass paragraphClass)
        {
            return paragraphClass switch
            {
                ParagraphClass.b1 => new HeadingFirstLevelGOST_7_0_11(),
                ParagraphClass.b2 => new HeadingOtherLevelsGOST_7_0_11(ParagraphClass.b2),
                ParagraphClass.b3 => new HeadingOtherLevelsGOST_7_0_11(ParagraphClass.b3),
                ParagraphClass.b4 => new HeadingOtherLevelsGOST_7_0_11(ParagraphClass.b4),
                ParagraphClass.c1 => new ParagraphRegularGOST_7_0_11(),
                ParagraphClass.c2 => new ParagraphBeforeListGOST_7_0_11(),
                ParagraphClass.c3 => new ParagraphBeforeEquationGOST_7_0_11(),
                ParagraphClass.d1 => new SimpleListFirstElementGOST_7_32(),
                ParagraphClass.d2 => new SimpleListMiddleElementGOST_7_32(),
                ParagraphClass.d3 => new SimpleListLastElementGOST_7_32(),
                ParagraphClass.d4 => new ComplexListFirstElementGOST_7_32(),
                ParagraphClass.d5 => new ComplexListMiddleElementGOST_7_32(),
                ParagraphClass.d6 => new ComplexListLastElementGOST_7_32(),
                //ParagraphClass.f1 => new TableSignGOST_7_32(ParagraphClass.f1),
                //ParagraphClass.f3 => new TableSignGOST_7_32(ParagraphClass.f3),
                //ParagraphClass.h1 => new ImageSignGOST_7_32(),
                //ParagraphClass.r0 => new SourcesListElementGOST_7_32(),
                _ => null
            };
        }
    }
}

﻿using System;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class Paragraph : DocumentElement
    {
        //c0
        public override string[] Suffixes => new string[] { ".", "!", "?" };
        
        // TODO: Внутри параграфа МОЖЕТ встречаться фраза, оформленная курсивом
    }
}
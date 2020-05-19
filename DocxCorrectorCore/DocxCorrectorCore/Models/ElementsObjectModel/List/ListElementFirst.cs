using System;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ListElementFirst : ListElement
    {
        //d1
        public override string[] Suffixes => new string[] {",", ":"};
        
        // TODO: Начинается с тире ИЛИ цифры 1 ИЛИ русской буквы "a"
        // TODO: Предыдущий параграф заканчивается на ":"
    }
}
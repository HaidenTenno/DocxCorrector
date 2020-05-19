using System;
using System.Text.RegularExpressions;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class ListElement : DocumentElement, IRegexSupportable
    {
        //d0
        // TODO: Каждый элемент перечисления начинается с тире (-) ИЛИ
        public override string[] Prefixes => new string[] { "-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" };
        //  TODO: ИЛИ строчной буквы, начиная с буквы "а" (за исключением букв ё, з, й, о, ч, ъ, ы, ь), ИЛИ арабской цифры, после которых ставится скобка
        public Regex Regex => throw new NotImplementedException();
        // TODO: Если элемент сделан НЕ средствами Word, то после маркера (любого вида), должен стоять пробел
    }
}

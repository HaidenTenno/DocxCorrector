using System;
using System.Text.RegularExpressions;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public class Heading //: DocumentElement
    {
        //b0
        // public override bool Bold => true;
        // public override int? EmptyLinesAfter => 1;
        
        // TODO: Параграф не может быть того же типа, что и предыдущий (у заголовка недопустимы переносы строк)
        // TODO: Заголовок НЕ может заканчиватся знаком препинания
        // TODO: После каждого заголовка идёт пустая строка ИЛИ нужный отсуп после
    }
}
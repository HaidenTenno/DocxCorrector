﻿using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class SourcesListElement : ListElement
    {
        // public override string[] Suffixes => new string[] {",", ":"};

        // TODO: Начинается с тире ИЛИ цифры 1 ИЛИ русской буквы "a"
        // TODO: Предыдущий параграф заканчивается на ":"

        public List<string> KeyWords = new List<string> { "Список литературы", "Список источников", "Список использованных", "Список использованной" };

        public override List<Regex> Regexes => new List<Regex>
        {
            new Regex(@"\d( [A-ZА-Я][a-zа-я]*\,? ([A-ZА-Я]\.){1,2}\,?){1,3} [A-ZА-Я].* (- ([A-ZА-Я]\. ?){1,3})?\: [A-ZА-Я][a-zа-я].*\, [1-9]\d{2,3}\. - [1-9]\d*( \- [1-9]\d*)? с\."),
            new Regex(@"\d ([А-ЯA-Z][a-zа-я]{1,} ([A-ZА-Я]\.\,?){1,2}){1,3} [А-ЯA-Z].* \/\/ ([А-ЯA-Z].*\: [А-ЯA-Z].*\/ ([а-яa-z].*\; [А-ЯA-Z].*(\. - [А-ЯA-Z].*\: )?\,) [1-9]\d{2,3}\. - С.[1-9]\d{1,}-\d*\.|[А-ЯA-Z].*\: [А-ЯA-Z].*\/ г.[А-ЯA-Z].*\, [(][а-я]* [1-9]\d{2,3} г\.[)]\.( - [A-ZА-Яа-яa-z]\.([1-9]\d*\.)?){0,4}\, [1-9]\d*\. - С.[1-9]\d*(\.|\-[1-9]\d*)\.|[A-ZА-Я].* \- [1-9]\d{2,3}\. - [A-Z] \d*\. - С\.[1-9]\d*(\-|\,)[1-9]\d*\.)"),
            new Regex(@"\d [A-ZА-Я].*\.(\: .*\. - [1,2]\d{3}\. | - URL: http[s]?:\/\/([a-z]*\.){1,})([a-zA-Z].[^а-яА-Я]*\/)(\([а-яА-Я ]*(0[1-9]|1[0-9]|2[0-9]|3[0-1])\.(0[1-9]|1[0-2])\.([1,2]\d{3}\)\.))?|[a-zA-Z].[^а-яА-Я]*[a-zа-яА-Я ]*(0[1-9]|1[0-9]|2[0-9]|3[0,1])\.(0[1-9]|1[0-2])\.([1,2]\d{3})\)\."),
            new Regex(@"\n\d{1,3}\.\ (([А-Я]([а-я]{1,})\ ([А-Я]([а-я]{0,1})\.\ ){1,2}))")
        };

        public SourcesListCorrections? CheckSourcesList(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            Word.Paragraph paragraph;
            try { paragraph = classifiedParagraphs[id].Paragraph; } catch { return null; }

            ParsedListElement parsedListElement = new ParsedListElement(paragraph);

            foreach (Regex regex in Regexes)
            {
                if (regex.IsMatch(parsedListElement.Content))
                {
                    return null;
                }
            }

            return new SourcesListCorrections(
                paragraphID: id,
                prefix: GemBoxHelper.GetParagraphPrefix(paragraph, 20),
                message: "Элемент списка литературы не соответствует ни одному из шаблонов"
            );
        }
    }
}
﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public sealed class ParsedListElement
    {
        public string Marker { get; }
        public string Body { get; }

        public string Content
        {
            get
            {
                return string.Join(" ", new string[] { Marker, Body });
            }
        }

        public ParsedListElement(Word.Paragraph paragraph)
        {
            if (paragraph.ListFormat.IsList)
            {
                Marker = paragraph.ListItem.ToString();
                Body = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
            }
            else
            {
                string content = GemBoxHelper.GetParagraphContentWithoutNewLine(paragraph);
                List<string> words = content.Split(' ').ToList();
                try { Marker = words[0]; } catch { Marker = ""; }
                words.RemoveAt(0);
                Body = string.Join(" ", words.ToArray());
            }
        }
    }

    public class ListElement : IRegexSupportable //: DocumentElement
    {
        //d0
        public virtual List<Regex> Regexes => throw new NotImplementedException();
    }
}

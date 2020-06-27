using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocxCorrectorCore.Models.Corrections;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class SourcesListElement //: DocumentElement
    {

        public SourcesListMistake? CheckSourcesListElement(int id, List<Regex> regexes, Word.Element sourcesListElement)
        {
            Word.Paragraph sourcesListElementParagraph;
            try { sourcesListElementParagraph = (Word.Paragraph)sourcesListElement; }
            catch
            {
                return new SourcesListMistake(
                    paragraphID: id,
                    prefix: "TABLE",
                    message: $"В списке литературы не может стоять таблица"
                );
            }

            ParsedListElement parsedListElement = new ParsedListElement(sourcesListElementParagraph);

            foreach (Regex regex in regexes)
            {
                if (regex.IsMatch(parsedListElement.Content))
                {
                    return null;
                }
            }

            return new SourcesListMistake(
                paragraphID: id,
                prefix: GemBoxHelper.GetParagraphPrefix(sourcesListElementParagraph, 20),
                message: "Элемент списка литературы не соответствует ни одному из шаблонов"
            );
        }
    }

    public class SourcesList
    {
        public List<string> KeyWords = new List<string> { "Список литературы", "Список источников", "Список использованных", "Список использованной" };

        public List<Regex> Regexes => new List<Regex>
        {
            // Статья в периодических изданиях и сборниках статей
            new Regex(@"\d+\.( [А-ЯA-Z][а-яa-z]*\,? ([A-ZА-Я]\.){1,2}\,?){1,3} [A-ZА-Я].* \/\/ [A-ZА-Я].*\. \- [1-9]\d{2,3}\. \- N [1-9]\d*\. - С\.( )?[1-9]\d*-[1-9]\d*\."),
            // Книги, монографии
            new Regex(@"^\d+\.(?> [А-ЯA-Z][а-яa-z-–]*,? (?>[A-ZА-Я]\.){1,2},?){0,3} [A-ZА-Я].*\/?\/? ?[A-ZА-Я].*\.(?> -| –)?(?> [А-ЯA-Z][^:;,]*(?>:[^:;,]+)?(?>;|,))* \d{2,4}\.(?>(?> -| –)? \d{1,2} [а-яa-z]+\.)?(?> -| –)?(?> Vol\. \d+,)?(?> N \d+(?> \(\d+\))?\.(?> -| –)?| No\. \d+\.)?(?> Т\. \d+(?>(?>-|–)\d+)?\.(?> -| –)?)?(?> (?>С|P)\. ?\d+? ?(?>-|–)? ?\d+\.| \d+ с\.)?(?> \(.*\)\.)?$"),
            // Тезисы докладов, материалы конференций
            new Regex(@"\d+\. ([А-ЯA-Z][a-zа-я]{1,} ([A-ZА-Я]\.\,?){1,2}){1,3} [А-ЯA-Z].* \/\/ ([А-ЯA-Z].*\: [А-ЯA-Z].*\/ ([а-яa-z].*\; [А-ЯA-Z].*(\. - [А-ЯA-Z].*\: )?\,) [1-9]\d{2,3}\. - С.[1-9]\d{1,}-\d*\.|[А-ЯA-Z].*\: [А-ЯA-Z].*\/ г.[А-ЯA-Z].*\, [(][а-я]* [1-9]\d{2,3} г\.[)]\.( - [A-ZА-Яа-яa-z]\.([1-9]\d*\.)?){0,4}\, [1-9]\d*\. - С.[1-9]\d*(\.|\-[1-9]\d*)\.|[A-ZА-Я].* \- [1-9]\d{2,3}\. - [A-Z] \d*\. - С\.[1-9]\d*(\-|\,)[1-9]\d*\.)"),
            // Патентная документация согласно стандарту ВОИС
            new Regex(@"\d+\.( [A-Z]{2,3} | [0-9]{3})[1-9]\d* [A-Z][1-9][,] (0[1-9]\.|1[0-2]\.]){2}[1,2]\d{3}\."),
            // Электронные ресурсы (обычные, с доступом, с диска)
            new Regex(@"^\d+\. [A-ZА-Я].* \/\/ .*\.(?> \[\d{4}\]\.| \d{4}\.) URL: https?:.* \(дата обращения: \d{2}(?>\.|-)\d{2}(?>\.|-)\d{4}\)\.$"),
            new Regex(@"^\d+\. [A-ZА-Я].*\[Электронный ресурс\]: .*Доступ из.*\.$"),
            new Regex(@"^\d+\. [A-ZА-Я].*,(?> \[\d{4}\]\.| \d{4}\.) \d+ электрон. опт. диск[а-я]* \([\w-–]+\)\."),
            // Нормативные документы (ГОСТ, ISO)
            new Regex(@"\d+\.? [А-Я]* [0-9.:*,""-]*( [А-Яа-я][а-я ]*.){1,} (- [А-Я].*\:)? ([А-Я][а-я]*.), [1,2]\d{3}. - [0-9]\d*( - [0-9]\d*)? (с|С)\."),
            new Regex(@"\d+\.? [A-Z]* \d*-\d{1,3}:[1,2]\d{3}\. [A-Z][a-z]* ([ A-Za-z ]*( )?[-.,:`:;?%'""0-9 ]){1,}URL: http(s)?:\/\/([a - z]*[./ _ ? 0 - 9]){1,} ((\([а - я]*[:]) (0[1-9]|1[0-9]|2[0-9]|3[0,1])\.(0[1-9]|1[0-2])\.([1, 2]\d{3})\))?\."),
            // Приказы правительства и пр.
            new Regex(@"\d\. ([А-Я][а-я ]*)* от ([1-9](\d{1})? [а-я]* [1,2]\d{3} г|([0-2][0-9][^00]|3[0,1])\.(0[1-9]|1[0-2])\.[1,2]\d{3} г)\. N(№)? [0-9]* ""[А-Я].*\""\.( - URL: http(s)?:\/\/([a-zA-Z]*[./_?0-9- ])*)?([(а-я: ]* ([0-2]\d{1}|3[0,2])\.(0\d{1}|1[0-2])\.([1,2]\d{3})\))?\."),
            // Книги (однотомные издания)
            new Regex(@"\d+\.? ([А-ЯA-Z][a-zа-я,`-]*(,)? ([A-ZА-Я].){1,2}|[A-ZА-Я][a-zа-я]* \([A-ZА-Я][a-zа-я]* ([A-ZА-Я]\.( )?){1,2}\)\.) [A-ZА-Я][a-zа-я].*\[[A-ZА-Я][a-zа-я]*\] (\/|:) [A-ZА-Яа-яa-z].*\. - [A-ZА-Я]*([а-яa-z-]*)?(.)? : [A-Za-zА-Яа-я].*\. - ((ISBN)? [0-9-:]*|N [0-9-:])?(\.| \([a-zа-я ,./:;0-9-]*\)\.)"),
            new Regex(@"\d+\.? ([A-ZА-Я][a-zа-я]*(,)? ([A-ZА-Я]\.( )?){1,2}){1,3} [A-ZА-Я0-9].*\[[A-ZА-Я][a-zа-я]*\] : [A-Za-zА-Яа-я.,;: -]* \/ (([A-ZА-Я][A-Zа-я]*(,) ([A-ZА-Я]\.( |, )){1,2}){1,3}|(([A-ZА-Я]\. ){1,2}[A-ZА-Я][a-zа-я]*(, )?){1,3}){1,3} \; [A-Zа-яa-zА-Я].* [1,2]\d{3}\.( - \d{1,}){1,} с\.(;|:).* - ISBN [0-9-]* ([a-zа-я .,()-]*)?\."),
            // Законодательные материалы (запись под заголовком, запись под заглавием)
            new Regex(@"\d+\.? ([ А-Я][а-я ]*)*\. [А-Я]([А-Яа-я0-9().,:; -]*)*\[[А-Я][а-я]*\] : [а-я.,0-9- ]* (- |: )([А-Я]\. : )?(.*\- ISBN [0-9-.]*[-А-Я]?\.)"),
            new Regex(@"\d+\.? ([А-Я][а-я ]*){1,}\[[А-Я][а-я]*\]( : |\. - [А-Я]\. : )(.*\. - ISBN [0-9-]*\.|\[.*\] : .* \/ .* - [А-Я] : .* - ISBN [0-9-]*\.)"),
            // Правила
            new Regex(@"\d+\.? ([А-Я][а-я ]*){1,}\[[А-Я][а-я]*\]( : |\. - [А-Я]\. : )(.*\. - ISBN [0-9-]*\.|\[.*\] : .* \/ .* - [А-Я] : .* - ISBN [0-9-]*\.)"),
            new Regex(@"\d+\.? ([А-Я][а-я() ]*){1,}\[[А-Я][а-я]*\] : [А-Я]{2} [0-9-.:]* [А-Яа-я0-9.;,: ()-]*(.*)?\ - ISBN [0-9-:А-Я]*\."),
            new Regex(@"\d+\.? [A-Z]* \d*-\d{1,3}:[1,2]\d{3}\. [A-Z][a-z]* ([ A-Za-z ]*( )?[-.,:`:;?%'""0-9 ]){1,}URL: http(s)?:\/\/([a - z]*[./ _ ? 0 - 9]){1,} ((\([а - я]*[:]) (0[1-9]|1[0-9]|2[0-9]|3[0,1])\.(0[1-9]|1[0-2])\.([1, 2]\d{3})\))?\."),
            // Многотомные издания (документ в целом)
            new Regex(@"\d+\.? ([А-Я][а-я]*(,) ([А-Я]. ){1,2}){1,3}[А-Я][а-я]* \[[А-Я][а-я]*\] : .* \/ [А-я ]*; \[[А-Яа-я(),./;: -]*\]\. - [А-Я]\. (: [А-Я][а-я, -]*){1,4}[1,2]\d{3}\..* - ISBN [0-9-]* .*\.(\nТ\. \d{1} : ([А-Я][а-я.]*(.)?(:)? )*- \d{1,}(-\d{1,})? с. - ([А-Я][а-я. ]*(.)?(:)? )*(с. \d{1,}-\d{1,})?((\. - |\.: |; )[А-Я][а-я ]*){1,}.* - ISBN [0-9-]*\.){1,}"),
            // Депонированные научные работы
            new Regex(@"\d*(\.)? [А-ЯA-Z][a-zа-я]*(,)? ([A-ZА-Я](\. )){1,}.* \[[A-ZА-Я][а-яa-z]*\] \/( ([A-ZА-Я](\.) ){1,2}[A-ZА-Я][a-zа-я]*(,)?){1,} ; [A-ZА-Я][A-ZА-Яa-zа-я.,/;:() -]*\d{4}\. - \d{1,} с\. : [A-ZА-Яa-zа-я.,/;:() -]*\d{1,}-\d{1,}\. - [A-ZА-Яa-zа-я.,/;:() -]* (([0-2][0-9][^00]|3[0-1])(0[1-9]|1[0-2])\.([0-9][0-9])\,)\1 N(№)? \d{1,}\.\n[A-ZА-Я][a-zа-я].*\[[A-ZА-Я][a-zа-я]*\] \/ ([A-ZА-Я]\. ){1,}[А-ЯA-Z][a-zа-я]* .* \; .*\. - [A-ZА-Я]\.\, [1,2]\d{3}\.- \d{3} с. - .* с\. \d{3}-\d{3}\. - .* \1\d{1,}\."),
            // Изоиздания
            new Regex(@"\d+\.? [A-ZА-Я][a-zа-я]*(,)? ([A-ZА-Я](\.) ){1,2}[A-ZА-Я][a-яa-z ]*(,)? [1,2]\d{3} \[[A-ZА-Я][a-zа-я]*\] : (.*)? \/ ([A-ZА-Я]\. ){1,2}[A-ZА-Я][a-zа-я]* \([1,2]\d{3}-[1,2]\d{3}\) ; [A-ZА-Я][A-ZА-Яa-zа-я.,/;:""'() -]*[1,2]\d{3}\. - (.*)\."),
            // Нотные издания
            new Regex(@"\d+\.? [A-ZА-Я][a-zа-я]*(,)? ([A-ZА-Я](\.) ){1,2}[A-ZА-Я][a-яa-z ]*(,)? \[[A-ZА-Я][a-zа-я]*\] : ((\(([A-ZА-Я][a-zа-я]*( )?){1,2})\) : )?[a-zа-я0-9.,A-ZА-Я:/; ]*(\[[a-zа-я0-9.,A-ZА-Я ]*\]\. )?- [A-ZА-Я][a-zа-я ]*\. - [A-ZА-Я]\. : [A-ZА-Я][a-zа-я ]*\, [1,2]\d{3}\.(.*)?\."),

        };

        public SourcesListCorrections? CheckSourcesList(int id, List<ClassifiedParagraph> classifiedParagraphs)
        {
            List<SourcesListMistake> sourcesListMistakes = new List<SourcesListMistake>();

            // Идем до конца документа ИЛИ пока не встретим следующий заголовок
            for (int sourcesListElementIndex = id + 1; sourcesListElementIndex < classifiedParagraphs.Count(); sourcesListElementIndex++)
            {
                if (classifiedParagraphs[sourcesListElementIndex].ParagraphClass == null) { continue; }

                if (classifiedParagraphs[sourcesListElementIndex].ParagraphClass == ParagraphClass.b1) { break; }

                var standartSourcesListElement = new SourcesListElement();

                // ПРОВЕРКА НАЧИНАЕТСЯ
                SourcesListMistake? currentSourcesListMistakes = standartSourcesListElement.CheckSourcesListElement(sourcesListElementIndex, Regexes, classifiedParagraphs[sourcesListElementIndex].Element);
                if (currentSourcesListMistakes != null) { sourcesListMistakes.Add(currentSourcesListMistakes); }
            }

            if (sourcesListMistakes.Count != 0)
            {
                return new SourcesListCorrections(
                    paragraphID: id,
                    prefix: GemBoxHelper.GetParagraphPrefix((Word.Paragraph)classifiedParagraphs[id].Element, 20),
                    mistakes: sourcesListMistakes
                );
            }
            else
            {
                return null;
            }
        }
    }
}
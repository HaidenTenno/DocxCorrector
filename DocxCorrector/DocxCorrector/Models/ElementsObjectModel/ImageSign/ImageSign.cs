using System;
using System.Text.RegularExpressions;

namespace DocxCorrector.Models.ElementsObjectModel
{
    public class ImageSign : DocumentElement, IRegexSupportable
    {
        public override AligmentType Aligment => AligmentType.Center;

        // IRegexSupportable
        public Regex Regex => new Regex(@"^Рисунок (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?$");
    }
}

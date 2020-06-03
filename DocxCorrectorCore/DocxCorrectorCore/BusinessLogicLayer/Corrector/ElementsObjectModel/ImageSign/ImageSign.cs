using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public class ImageSign : IRegexSupportable //: DocumentElement, IRegexSupportable
    {
        // public override AlignmentType Alignment => AlignmentType.Center;

        // IRegexSupportable
        public List<Regex> Regexes => new List<Regex> 
        { 
            new Regex (@"^Рисунок (?>[А-ЕЖИК-НП-ЦШЩЭЮЯ]\.[\d]+|[\d]+(?>\.[\d]+)?)(?> - .*)?$") 
        };
    }
}

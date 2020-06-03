using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DocxCorrectorCore.BusinessLogicLayer.Corrector.ElementsObjectModel
{
    public interface IRegexSupportable
    {
        public List<Regex> Regexes { get; }
    }
}

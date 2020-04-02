using System;
using System.Text.RegularExpressions;

namespace DocxCorrector.Models.ElementsObjectModel
{
    public class ListElement : DocumentElement, IRegexSupportable
    {
        public override string[] Prefixes => new string[] { "-", "־", "᠆", "‐", "‑", "‒", "–", "—", "―", "﹘", "﹣", "－" };

        public override string[] Suffixes => new string[] { ";", "," };

        // IRegexSupportable
        public Regex Regex => throw new NotImplementedException();
    }
}

using System;
using System.Text.RegularExpressions;

namespace DocxCorrector.Models.ElementsObjectModel
{
    public interface IRegexSupportable
    {
        public Regex Regex { get; }
    }
}

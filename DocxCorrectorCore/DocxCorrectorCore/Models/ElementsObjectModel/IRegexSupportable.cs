using System;
using System.Text.RegularExpressions;

namespace DocxCorrectorCore.Models.ElementsObjectModel
{
    public interface IRegexSupportable
    {
        public Regex Regex { get; }
    }
}

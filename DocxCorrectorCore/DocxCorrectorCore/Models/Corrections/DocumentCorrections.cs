using System;
using System.Collections.Generic;

namespace DocxCorrectorCore.Models
{
    // TODO: Find another place
    public enum RulesModel
    {
        GOST,
        ITMO
    }

    public sealed class DocumentCorrections
    {
        public readonly RulesModel RulesModel;
        public readonly List<ParagraphCorrections> ParagraphsCorrections;

        public DocumentCorrections(RulesModel rules, List<ParagraphCorrections> paragraphsCorrections)
        {
            RulesModel = rules;
            ParagraphsCorrections = paragraphsCorrections;
        }

        public DocumentCorrections(RulesModel rules)
        {
            RulesModel = rules;
            ParagraphsCorrections = new List<ParagraphCorrections>();
        }
    }
}

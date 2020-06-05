using System.Collections.Generic;

namespace DocxCorrectorCore.Models.Corrections
{
    public enum RulesModel
    {
        GOST,
        ITMO
    }

    public sealed class DocumentCorrections
    {
        public readonly RulesModel RulesModel;
        public readonly List<ParagraphCorrections> ParagraphsCorrections;
        public readonly List<SourcesListCorrections> SourcesListCorrections;
        public readonly List<TableCorrections> TablesCorrections;

        public DocumentCorrections(
            RulesModel rules,
            List<ParagraphCorrections> paragraphsCorrections,
            List<SourcesListCorrections> sourcesListCorrections,
            List<TableCorrections> tablesCorrections
        )
        {
            RulesModel = rules;
            ParagraphsCorrections = paragraphsCorrections;
            SourcesListCorrections = sourcesListCorrections;
            TablesCorrections = tablesCorrections;
        }

        public DocumentCorrections(RulesModel rules)
        {
            RulesModel = rules;
            ParagraphsCorrections = new List<ParagraphCorrections>();
            SourcesListCorrections = new List<SourcesListCorrections>();
            TablesCorrections = new List<TableCorrections>();
        }
    }
}

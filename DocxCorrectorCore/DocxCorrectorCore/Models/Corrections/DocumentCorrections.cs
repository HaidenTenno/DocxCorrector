using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace DocxCorrectorCore.Models.Corrections
{
    public enum RulesModel
    {
        GOST,
        ITMO
    }

    [JsonConverter(typeof(StringEnumConverter))]
    public enum MistakeImportance
    {
        Warning,
        Regular,
        Critical
    }

    public sealed class DocumentCorrections
    {
        public readonly RulesModel RulesModel;
        public readonly List<ParagraphCorrections> ParagraphsCorrections;
        public readonly List<SourcesListCorrections> SourcesListCorrections;
        public readonly List<TableCorrections> TablesCorrections;
        public readonly List<HeadlingCorrections> HeadlingCorrections;

        public DocumentCorrections(
            RulesModel rules,
            List<ParagraphCorrections> paragraphsCorrections,
            List<SourcesListCorrections> sourcesListCorrections,
            List<TableCorrections> tablesCorrections,
            List<HeadlingCorrections> headlingCorrections
        )
        {
            RulesModel = rules;
            ParagraphsCorrections = paragraphsCorrections;
            SourcesListCorrections = sourcesListCorrections;
            TablesCorrections = tablesCorrections;
            HeadlingCorrections = headlingCorrections;
        }

        public DocumentCorrections(RulesModel rules)
        {
            RulesModel = rules;
            ParagraphsCorrections = new List<ParagraphCorrections>();
            SourcesListCorrections = new List<SourcesListCorrections>();
            TablesCorrections = new List<TableCorrections>();
            HeadlingCorrections = new List<HeadlingCorrections>();
        }
    }
}

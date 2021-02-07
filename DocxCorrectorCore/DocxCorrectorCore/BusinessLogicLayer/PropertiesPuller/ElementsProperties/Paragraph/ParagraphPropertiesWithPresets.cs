using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models.Corrections;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public class ParagraphPropertiesWithPresets : ParagraphPropertiesTableZero
    {
        // Public
        public List<ParagraphClass> ClassesFromRequirements { get; }
        public List<ParagraphClass> ClassesFromWord { get; }
        public List<ParagraphClass> ClassesFromUser { get; }

        public ParagraphPropertiesWithPresets(int id, Word.Paragraph paragraph, CombinedPresetValues combinedPresetValues) : base(id, paragraph)
        {
            ClassesFromRequirements = combinedPresetValues.GetSimilarClasses(CombinedPresetValues.PresetSource.Requirements, paragraph);
            ClassesFromWord = combinedPresetValues.GetSimilarClasses(CombinedPresetValues.PresetSource.Word, paragraph);
            ClassesFromUser = combinedPresetValues.GetSimilarClasses(CombinedPresetValues.PresetSource.User, paragraph);
        }

        public ParagraphPropertiesWithPresets(int id, string content) : base(id, content) 
        {
            ClassesFromRequirements = new List<ParagraphClass>();
            ClassesFromWord = new List<ParagraphClass>();
            ClassesFromUser = new List<ParagraphClass>();
        }
    }
}

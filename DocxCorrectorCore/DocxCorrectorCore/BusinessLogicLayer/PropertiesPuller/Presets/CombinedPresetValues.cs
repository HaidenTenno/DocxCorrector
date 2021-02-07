using System;
using DocxCorrectorCore.Models.Corrections;
using System.Collections.Generic;
using Word = GemBox.Document;

namespace DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller
{
    public sealed class CombinedPresetValues
    {
        public enum PresetSource
        {
            Requirements,
            Word,
            User
        }

        public readonly List<PresetValue> FromRequirements;
        public readonly List<PresetValue> FromWord;
        public readonly List<PresetValue> FromUser;

        public CombinedPresetValues(List<PresetValue> fromRequirements, List<PresetValue> fromWord, List<PresetValue> fromUser)
        {
            FromRequirements = fromRequirements;
            FromWord = fromWord;
            FromUser = fromUser;
        }

        // Получить список классов из пресетов, на которые похож параграф
        public List<ParagraphClass> GetSimilarClasses(PresetSource presetSource, Word.Paragraph paragraph)
        {
            List<PresetValue> selectedPresetValues = new List<PresetValue>();

            switch (presetSource)
            {
                case PresetSource.Requirements:
                    selectedPresetValues = FromRequirements;
                    break;
                case PresetSource.User:
                    selectedPresetValues = FromUser;
                    break;
                case PresetSource.Word:
                    selectedPresetValues = FromWord;
                    break;
            }

            List<ParagraphClass> similarParagraphClasses = new List<ParagraphClass>();

            foreach (PresetValue presetValue in selectedPresetValues)
            {
                if (presetValue.ParagraphLooksLikePreset(paragraph)) { similarParagraphClasses.Add(presetValue.ParagraphClass); }
            }

            return similarParagraphClasses;
        }
    }
}

﻿using System;
using DocxCorrectorCore.App;
using DocxCorrectorCore.Models.Corrections;

// TODO: Remove
namespace DocxCorrectorCore.UserDialog
{
    public sealed class FixDocumentQuestionController : StringAnswerQuestionController
    {
        // Public
        public FixDocumentQuestionController() : base("Введите: \nПуть к файлу для проверки, \nТребования для проверки (GOST/ITMO), \nПуть к JSON файлу с классами параграфов, \nПуть к файлу или директории для сохранения результата") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(4)) { return; }

            RulesModel chosenRules;
            try
            {
                chosenRules = (RulesModel)Enum.Parse(typeof(RulesModel), UserAnswer[1]);
            }
            catch
            {
                Console.WriteLine("Выбраны ошибочные требования");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateFixedDocument(fileToFix: UserAnswer[0], rules: chosenRules, paragraphsClassesFile: UserAnswer[2], resultPath: UserAnswer[3]);
        }
    }
}
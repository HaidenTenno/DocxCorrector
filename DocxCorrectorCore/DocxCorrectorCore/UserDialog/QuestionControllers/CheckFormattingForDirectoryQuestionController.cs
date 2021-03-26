using System;
using DocxCorrectorCore.App;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class CheckFormattingForDirectoryQuestionController : StringAnswerQuestionController
    {
        // Public
        public CheckFormattingForDirectoryQuestionController() : base("Введите: \nПуть к корневой директории, \nТребования для проверки (GOST/ITMO)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(2)) { return; }

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
            featuresProvider.GenerateParagraphMistakesFiles(UserAnswer[0], chosenRules);
        }
    }
}
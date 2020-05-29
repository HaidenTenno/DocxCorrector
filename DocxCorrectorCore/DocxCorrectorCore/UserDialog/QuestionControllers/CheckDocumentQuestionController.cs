using DocxCorrectorCore.App;
using System;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class CheckDocumentQuestionController : StringAnswerQuestionController
    {
        // Public
        public CheckDocumentQuestionController() : base("Введите: \nПуть к файлу для проверки, \nТребования для проверки (GOST/ITMO), \nПуть к JSON файлу с классами параграфов, \nПуть к файлу или директории для сохранения результата") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(4)) { return; }

            Models.RulesModel chosenRules;
            try
            {
                chosenRules = (Models.RulesModel)Enum.Parse(typeof(Models.RulesModel), UserAnswer[1]);
            }
            catch
            {
                Console.WriteLine("Выбраны ошибочные требования");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateMistakesJSON(fileToCorrect: UserAnswer[0], rules: chosenRules, paragraphsClassesFile: UserAnswer[2], resultPath: UserAnswer[3]);
        }
    }
}


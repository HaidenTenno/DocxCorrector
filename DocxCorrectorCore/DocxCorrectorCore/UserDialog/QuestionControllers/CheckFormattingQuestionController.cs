using System;
using DocxCorrectorCore.App;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class CheckFormattingQuestionController : StringAnswerQuestionController
    {
        // Public
        public CheckFormattingQuestionController() : base("Введите: \nПуть к файлу для проверки, \nТребования для проверки (GOST/ITMO), \nНомер параграфа, \nМетка класса выбранного параграфа, \nПуть к файлу или директории для сохранения списка ошибок") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(5)) { return; }

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

            if (!int.TryParse(UserAnswer[2], out int chosenParagraphID))
            {
                Console.WriteLine("Номер параграфа должен быть числом");
                return;
            }

            ParagraphClass chosenClass;
            try
            {
                chosenClass = (ParagraphClass)Enum.Parse(typeof(ParagraphClass), UserAnswer[3]);
            }
            catch
            {
                Console.WriteLine("Выбран ошибочный класс");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateFormattingMistakesJSON(fileToCorrect: UserAnswer[0], rules: chosenRules, paragraphID: chosenParagraphID, paragraphClass: chosenClass, resultPath: UserAnswer[4]);
        }
    }
}
using System;
using DocxCorrectorCore.App;
using DocxCorrectorCore.Models.Corrections;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class CreateClassModelFileQuestionController : StringAnswerQuestionController
    {
        // Public
        public CreateClassModelFileQuestionController() : base("Введите: \nТребования для проверки (GOST/ITMO), \nМетка класса параграфа, \nПуть к файлу или директории для сохранения модели") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(3)) { return; }

            RulesModel chosenRules;
            try
            {
                chosenRules = (RulesModel)Enum.Parse(typeof(RulesModel), UserAnswer[0]);
            }
            catch
            {
                Console.WriteLine("Выбраны ошибочные требования");
                return;
            }

            ParagraphClass chosenClass;
            try
            {
                chosenClass = (ParagraphClass)Enum.Parse(typeof(ParagraphClass), UserAnswer[1]);
            }
            catch
            {
                Console.WriteLine("Выбран ошибочный класс");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateModelJSON(rules: chosenRules, paragraphClass: chosenClass, resultPath: UserAnswer[2]);
        }
    }
}
using System;
using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class TwoCSVsQuestionController : StringAnswerQuestionController
    {
        // Public
        public TwoCSVsQuestionController() : base("Введите: \nПуть к документу, \nНомер первого параграфа, \nПуть к файлу или директории для сохранения первой csv, \nПуть к файлу или директории для сохранения второй csv") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(4)) { return; }

            if (!int.TryParse(UserAnswer[1], out int chosenParagraphID))
            {
                Console.WriteLine("Номер параграфа должен быть числом");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateParagraphsPropertiesForAllTables(UserAnswer[0], chosenParagraphID, UserAnswer[2], UserAnswer[3]);
        }
    }
}
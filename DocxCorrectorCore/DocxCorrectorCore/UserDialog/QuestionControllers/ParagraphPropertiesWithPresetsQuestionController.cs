using System;
using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ParagraphPropertiesWithPresetsQuestionController : StringAnswerQuestionController
    {
        // Public
        public ParagraphPropertiesWithPresetsQuestionController() : base("Введите: \nПуть к документу,  \nПуть к JSON файлу с информацией из пресетов, \nНомер первого параграфа, \nПуть к файлу или директории для сохранения CSV файла со свойствами параграфов (с проставленными классами из пресетов)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(4)) { return; }

            if (!int.TryParse(UserAnswer[2], out int chosenParagraphID))
            {
                Console.WriteLine("Номер параграфа должен быть числом");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateCSVWithPresetsInfo(UserAnswer[0], UserAnswer[1], chosenParagraphID, UserAnswer[3]);
        }
    }
}
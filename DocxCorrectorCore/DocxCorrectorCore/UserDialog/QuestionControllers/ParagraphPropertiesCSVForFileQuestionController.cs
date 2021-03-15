using System;
using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ParagraphPropertiesCSVForFileQuestionController : StringAnswerQuestionController
    {
        // Public
        public ParagraphPropertiesCSVForFileQuestionController() : base("Введите: \nПуть к документу, \nНомер первого параграфа, \nПуть к файлу или директории для сохранения CSV файла со свойствами параграфов") { }

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
            featuresProvider.GenerateParagraphsPropertiesCSV(UserAnswer[0], chosenParagraphID, UserAnswer[2]);
        }
    }
}


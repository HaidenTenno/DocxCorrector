using System;
using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class CreatePresetInfoExampleQuestionController : StringAnswerQuestionController
    {
        // Public
        public CreatePresetInfoExampleQuestionController() : base("Введите: \nПуть к документу, \nНомер параграфа, \nПуть к файлу или директории для сохранения JSON файла со свойствами параграфа (пример пресета)") { }

        public override void Load()
        {
            //public void GeneratePresetInfoFromParagraph(string filePath, int paragraphID, string resultPath)


            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(3)) { return; }

            if (!int.TryParse(UserAnswer[1], out int chosenParagraphID))
            {
                Console.WriteLine("Номер параграфа должен быть числом");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GeneratePresetInfoFromParagraph(UserAnswer[0], chosenParagraphID, UserAnswer[2]);
        }
    }
}
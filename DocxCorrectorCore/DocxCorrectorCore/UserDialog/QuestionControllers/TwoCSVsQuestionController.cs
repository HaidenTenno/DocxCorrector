using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class TwoCSVsQuestionController : StringAnswerQuestionController
    {
        // Public
        public TwoCSVsQuestionController() : base("Введите: \nПуть к документу \nПуть к файлу или директории для сохранения CSV файла со свойствами параграфов \nПуть к файлу или директории для сохранения CSV файла со свойствами параграфов (для таблицы 0)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(3)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateParagraphsPropertiesForAllTables(UserAnswer[0], UserAnswer[1], UserAnswer[2]);
        }
    }
}


using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ParagraphPropertiesWithPresetsQuestionController : StringAnswerQuestionController
    {
        // Public
        public ParagraphPropertiesWithPresetsQuestionController() : base("Введите: \nПуть к документу,  \nПуть к JSON файлу с информацией из пресетов, \nПуть к файлу или директории для сохранения CSV файла со свойствами параграфов (с проставленными классами из пресетов)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(3)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateCSVWithPresetsInfo(UserAnswer[0], UserAnswer[1], UserAnswer[2]);
        }
    }
}
using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ParagraphPropertiesCSVForFileQuestionController : StringAnswerQuestionController
    {
        // Public
        public ParagraphPropertiesCSVForFileQuestionController() : base("Введите: \nПуть к документу \nПуть к директории для сохранения CSV файла со свойствами параграфов") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(2)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateParagraphsPropertiesCSV(UserAnswer[0], UserAnswer[1]);
        }
    }
}


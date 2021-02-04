using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class SectionPropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public SectionPropertiesJSONQuestionController() : base("Введите: \nПуть к документу, \nПуть к директории для сохранения JSON файла со свойствами секций") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(2)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateSectionsPropertiesJSON(UserAnswer[0], UserAnswer[1]);
        }
    }
}


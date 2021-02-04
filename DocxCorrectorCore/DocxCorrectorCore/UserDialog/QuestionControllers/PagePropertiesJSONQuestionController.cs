using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class PagePropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public PagePropertiesJSONQuestionController() : base("Введите: \nПуть к документу, \nПуть к директории для сохранения JSON файла со свойствами страниц") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(2)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GeneratePagesPropertiesJSON(UserAnswer[0], UserAnswer[1]);
        }
    }
}


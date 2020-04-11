using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ParagraphPropertiesCSVQuestionController : StringAnswerQuestionController
    {
        // Public
        public ParagraphPropertiesCSVQuestionController() : base("Введите: \nПуть к корневой директории, в поддиректориях которой находятся документы, для которых нужно создать CSV со свойствами параграфов") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateCSVFiles(UserAnswer[0]);
        }
    }
}


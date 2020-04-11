using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class NormalizedParagraphPropertiesCSVQuestionController : StringAnswerQuestionController
    {
        // Public
        public NormalizedParagraphPropertiesCSVQuestionController() : base("Введите: \nПуть к корневой директории, в поддиректориях которой находятся документы, для которых нужно создать CSV с нормализованными свойствами параграфов") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateNormalizedCSVFiles(UserAnswer[0]);
        }
    }
}


using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class TableOfContentsInfoQuestionController : StringAnswerQuestionController
    {
        // Public
        public TableOfContentsInfoQuestionController() : base("Введите: \nПуть к документу, информацию о содержании которого нужно вывести в консоль") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.PrintTableOfContentsInfo(UserAnswer[0]);
        }
    }
}


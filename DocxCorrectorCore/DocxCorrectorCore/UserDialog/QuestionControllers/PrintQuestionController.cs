using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class PrintQuestionController : StringAnswerQuestionController
    {
        // Public
        public PrintQuestionController() : base("Введите: \nПуть к документу, параграфы которого нужно вывести в консоль") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.PrintParagraphs(UserAnswer[0]);
        }
    }
}


using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class StructureInfoQuestionController : StringAnswerQuestionController
    {
        // Public
        public StructureInfoQuestionController() : base("Введите: \nПуть к документу, информацию о структуре которого нужно вывести в консоль") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.PrintStructureInfo(UserAnswer[0]);
        }
    }
}


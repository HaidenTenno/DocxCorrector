using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class TestParagraphPropertiesPullingSpeedQuestionController : StringAnswerQuestionController
    {
        // Public
        public TestParagraphPropertiesPullingSpeedQuestionController() : base("Введите: \nПуть к корневой директории, в поддиректориях которой находятся документы, для которых нужно создать CSV со свойствами параграфов (будет произведен тест скорости методов)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.TestParagraphPropertiesPullingSpeed(UserAnswer[0]);
        }
    }
}

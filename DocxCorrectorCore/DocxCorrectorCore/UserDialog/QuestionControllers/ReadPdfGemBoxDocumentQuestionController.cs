using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ReadPdfGemBoxDocumentQuestionController : StringAnswerQuestionController
    {
        // Public
        public ReadPdfGemBoxDocumentQuestionController() : base("Введите: \nПуть к pdf файлу, содержимое которое необходимо вывести в консоль (Библиотека Gembox.Document)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.PrintPdfGemBoxDocument(UserAnswer[0]);
        }
    }
}


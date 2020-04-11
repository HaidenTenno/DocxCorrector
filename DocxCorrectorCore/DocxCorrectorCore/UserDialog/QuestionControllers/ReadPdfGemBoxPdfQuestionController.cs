using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class ReadPdfGemBoxPdfQuestionController : StringAnswerQuestionController
    {
        // Public
        public ReadPdfGemBoxPdfQuestionController() : base("Введите: \nПуть к pdf файлу, содержимое которое необходимо вывести в консоль (Библиотека Gembox.Pdf)") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(1)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.PrintPdfGemBoxPdf(UserAnswer[0]);
        }
    }
}
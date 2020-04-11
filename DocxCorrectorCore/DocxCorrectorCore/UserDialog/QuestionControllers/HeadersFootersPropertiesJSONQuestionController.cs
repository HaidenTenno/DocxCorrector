using System;
using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class HeadersFootersPropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public HeadersFootersPropertiesJSONQuestionController() : base("Введите: \nТип колонтитулов (0: верхний, 1: нижний) \nПуть к документу \nПуть к директории для сохранения JSON файла со свойствами колонтитулов") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(3)) { return; }

            Models.HeaderFooterType? chosenHeaderFooterType;
            switch (UserAnswer[0])
            {
                case "0":
                    chosenHeaderFooterType = Models.HeaderFooterType.Header;
                    break;
                case "1":
                    chosenHeaderFooterType = Models.HeaderFooterType.Footer;
                    break;
                default:
                    Console.WriteLine("Выбрана некорректная опция");
                    return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();

            featuresProvider.GenerateHeadersFootersInfoJSON((Models.HeaderFooterType)chosenHeaderFooterType, UserAnswer[1], UserAnswer[2]);
        }
    }
}


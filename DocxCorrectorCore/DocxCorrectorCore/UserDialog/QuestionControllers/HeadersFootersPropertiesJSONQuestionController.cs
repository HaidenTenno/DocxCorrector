using System;
using DocxCorrectorCore.App;
using DocxCorrectorCore.BusinessLogicLayer.PropertiesPuller;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class HeadersFootersPropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public HeadersFootersPropertiesJSONQuestionController() : base("Введите: \nТип колонтитулов (0: верхний, 1: нижний), \nПуть к документу, \nПуть к директории для сохранения JSON файла со свойствами колонтитулов") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(3)) { return; }

            HeaderFooterType? chosenHeaderFooterType;
            switch (UserAnswer[0])
            {
                case "0":
                    chosenHeaderFooterType = HeaderFooterType.Header;
                    break;
                case "1":
                    chosenHeaderFooterType = HeaderFooterType.Footer;
                    break;
                default:
                    Console.WriteLine("Выбрана некорректная опция");
                    return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();

            featuresProvider.GenerateHeadersFootersInfoJSON((HeaderFooterType)chosenHeaderFooterType, UserAnswer[1], UserAnswer[2]);
        }
    }
}


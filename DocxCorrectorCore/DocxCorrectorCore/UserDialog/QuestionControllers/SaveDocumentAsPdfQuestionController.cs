﻿using DocxCorrectorCore.App;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class SaveDocumentAsPdfQuestionController : StringAnswerQuestionController
    {
        // Public
        public SaveDocumentAsPdfQuestionController() : base("Введите: \nПуть к docx файлу, который необходимо сохранить как pdf, \nПуть к директории для сохранения результата") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (CheckIfWrongArgumentsCountPassed(2)) { return; }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.SaveDocumentAsPdf(UserAnswer[0], UserAnswer[1]);
        }
    }
}


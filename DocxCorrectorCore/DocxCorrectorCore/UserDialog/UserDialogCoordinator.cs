using System;
using System.Collections.Generic;

namespace DocxCorrectorCore.UserDialog
{
    public enum QuestionControllerType
    {
        Print,
        PageProperties,
        SectionProperties,
        HeadersFooters,
        ParagraphProperties,
        NormalizedParagraphProperties,
        SaveDocumentAsPdf,
        SavePagesAsPdf
    }

    public sealed class UserDialogCoordinator
    {
        // Private

        private readonly QuestionsNavigationController NavigationController;

        // Creators
        private IntAnswerQuestionController CreateMainMenu()
        {
            IntAnswerQuestionController mainMenu = new IntAnswerQuestionController(
                actions: new List<(string info, Action action)>()
                {
                    ("Печать всех параграфов в консоль", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.Print))),
                    ("Печать свойства странц в файл", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.PageProperties))),
                    ("Печать свойства секций в файл", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.SectionProperties))),
                    ("Печать свойств верхних / нижних колонтитулов", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.HeadersFooters))),
                    ("Генерация CSV для свойств параграфов", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.ParagraphProperties))),
                    ("Генерация CSV для нормализованных свойств параграфов", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.NormalizedParagraphProperties))),
                    ("Сохранение документа как pdf", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.SaveDocumentAsPdf))),
                    ("Сохранение каждой страницы документа отдельным pdf", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.SavePagesAsPdf)))
                }
            );
            return mainMenu;
        }

        private StringAnswerQuestionController CreateStringAnswerQC(QuestionControllerType type)
        {
            return type switch
            {
                QuestionControllerType.Print => new PrintQuestionController(),
                QuestionControllerType.PageProperties => new PagePropertiesJSONQuestionController(),
                QuestionControllerType.SectionProperties => new SectionPropertiesJSONQuestionController(),
                QuestionControllerType.HeadersFooters => new HeadersFootersPropertiesJSONQuestionController(),
                QuestionControllerType.ParagraphProperties => new ParagraphPropertiesCSVQuestionController(),
                QuestionControllerType.NormalizedParagraphProperties => new NormalizedParagraphPropertiesCSVQuestionController(),
                QuestionControllerType.SaveDocumentAsPdf => new SaveDocumentAsPdfQuestionController(),
                QuestionControllerType.SavePagesAsPdf => new SavePagesAsPdfQuestionController(),
                _ => throw new NotImplementedException()
            };
        }

        // Public
        public UserDialogCoordinator(QuestionsNavigationController navigationController)
        {
            NavigationController = navigationController;
        }

        public void Start()
        {
            NavigationController.PushQuestionController(CreateMainMenu());
            NavigationController.Run();
        }
    }
}
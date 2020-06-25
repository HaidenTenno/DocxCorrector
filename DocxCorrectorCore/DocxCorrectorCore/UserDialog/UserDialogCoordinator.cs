using System;
using System.Collections.Generic;

namespace DocxCorrectorCore.UserDialog
{
    // TODO: NOT SUPPORTED IN OUR DLL
    public enum QuestionControllerType
    {
        Print,
        StructureInfo,
        TableOfContentsInfo,
        //PageProperties,
        SectionProperties,
        HeadersFooters,
        ParagraphPropertiesForFile,
        ParagraphProperties,
        TestPropertiesPullingSpeed,
        //SaveDocumentAsPdf,
        //SavePagesAsPdf,
        //ReadPdfGemboxDocument,
        //ReadPdfGemboxPdf,
        TwoCSVs,
        CheckDocument
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
                    ("Информация о структуре документа", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.StructureInfo))),
                    ("Информация о содержании документа", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.TableOfContentsInfo))),
                    //("Печать свойства странц в файл", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.PageProperties))),
                    ("Печать свойства секций в файл", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.SectionProperties))),
                    ("Печать свойств верхних / нижних колонтитулов", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.HeadersFooters))),
                    ("Генерация CSV для свойств параграфов (один файл)", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.ParagraphPropertiesForFile))),
                    ("Генерация CSV для свойств параграфов (для директории)", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.ParagraphProperties))),
                    ("Тестирование скорости синхронных/асинхронных методов при вытягивании свойств параграфов", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.TestPropertiesPullingSpeed))),
                    //("Сохранение документа как pdf", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.SaveDocumentAsPdf))),
                    //("Сохранение каждой страницы документа отдельным pdf", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.SavePagesAsPdf))),
                    //("Чтение pdf документа библиотекой Gembox.Document", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.ReadPdfGemboxDocument))),
                    //("Чтение pdf документа библиотекой Gembox.Pdf", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.ReadPdfGemboxPdf))),
                    ("Генерация CSV для свойств параграфов (один файл) + CSV для таблицы 0", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.TwoCSVs))),
                    ("Проверить оформление docx документа", () => NavigationController.PushQuestionController(CreateStringAnswerQC(QuestionControllerType.CheckDocument)))
                }
            );
            return mainMenu;
        }

        private StringAnswerQuestionController CreateStringAnswerQC(QuestionControllerType type)
        {
            return type switch
            {
                QuestionControllerType.Print => new PrintQuestionController(),
                QuestionControllerType.StructureInfo => new StructureInfoQuestionController(),
                QuestionControllerType.TableOfContentsInfo => new TableOfContentsInfoQuestionController(),
                //QuestionControllerType.PageProperties => new PagePropertiesJSONQuestionController(),
                QuestionControllerType.SectionProperties => new SectionPropertiesJSONQuestionController(),
                QuestionControllerType.HeadersFooters => new HeadersFootersPropertiesJSONQuestionController(),
                QuestionControllerType.ParagraphPropertiesForFile => new ParagraphPropertiesCSVForFileQuestionController(),
                QuestionControllerType.ParagraphProperties => new ParagraphPropertiesCSVQuestionController(),
                QuestionControllerType.TestPropertiesPullingSpeed => new TestParagraphPropertiesPullingSpeedQuestionController(),
                //QuestionControllerType.SaveDocumentAsPdf => new SaveDocumentAsPdfQuestionController(),
                //QuestionControllerType.SavePagesAsPdf => new SavePagesAsPdfQuestionController(),
                //QuestionControllerType.ReadPdfGemboxDocument => new ReadPdfGemBoxDocumentQuestionController(),
                //QuestionControllerType.ReadPdfGemboxPdf => new ReadPdfGemBoxPdfQuestionController(),
                QuestionControllerType.TwoCSVs => new TwoCSVsQuestionController(),
                QuestionControllerType.CheckDocument => new CheckDocumentQuestionController(),
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
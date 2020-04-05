using System;
using System.Collections.Generic;

namespace DocxCorrectorCore.App
{
    public sealed class UserDialogCoordinator
    {
        // Private
        private bool ShouldEnd = false;

        private Stack<UserQuestion> QuestionStack;

        private void PushQuestion(UserQuestion question)
        {
            QuestionStack.Push(question);
        }

        private void PopQuestion()
        {
            if (QuestionStack.Count == 0) { return; }
            QuestionStack.Pop();
        }

        private void PopAll()
        {
            QuestionStack.Clear();
        }

        private void LoadTop()
        {
            if (QuestionStack.Count == 0)
            {
                ShouldEnd = true;
            }
            else
            {
                QuestionStack.Peek().Load();
            }
        }

        private void EndProgram()
        {
            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }

        // Creators
        private IntAnswerQuestion createMainMenu()
        {
            IntAnswerQuestion mainMenu = new IntAnswerQuestion(
                actions: new List<(string info, Action action)>()
                {
                    ("Печать всех параграфов в консоль", () => { PushQuestion(createStringAnswerQuesion(UserQuestionType.Print)); }),
                    ("Печать свойства странц в файл", () => { PushQuestion(createStringAnswerQuesion(UserQuestionType.PageProperties)); }),
                    ("Печать свойства секций в файл", () => { PushQuestion(createStringAnswerQuesion(UserQuestionType.SectionProperties)); }),
                    ("Печать свойств верхних / нижних колонтитулов", () => { PushQuestion(createStringAnswerQuesion(UserQuestionType.HeadersFooters)); }),
                    ("Генерация CSV для свойств параграфов", () => { PushQuestion(createStringAnswerQuesion(UserQuestionType.ParagraphProperties)); }),
                    ("Генерация CSV для нормализованных свойств параграфов", () => { PushQuestion(createStringAnswerQuesion(UserQuestionType.NormalizedParagraphProperties)); }),
                    ("Выход", () => PopAll())
                }
            );
            return mainMenu;
        }

        private StringAnswerQuestion createStringAnswerQuesion(UserQuestionType type)
        {
            return type switch
            {
                UserQuestionType.Print => new PrintQuestion(
                    onBack: () => PopQuestion(),
                    onExit: () => PopAll()
                ),
                UserQuestionType.PageProperties => new PagePropertiesJSONQuestion(
                    onBack: () => PopQuestion(),
                    onExit: () => PopAll()
                ),
                UserQuestionType.SectionProperties => new SectionPropertiesJSONQuestion(
                    onBack: () => PopQuestion(),
                    onExit: () => PopAll()
                ),
                UserQuestionType.HeadersFooters => new HeadersFootersPropertiesJSONQuestion(
                    onBack: () => PopQuestion(),
                    onExit: () => PopAll()
                ),
                UserQuestionType.ParagraphProperties => new ParagraphPropertiesCSVQuestion(
                    onBack: () => PopQuestion(),
                    onExit: () => PopAll()
                ),
                UserQuestionType.NormalizedParagraphProperties => new NormalizedParagraphPropertiesCSVQuestion(
                    onBack: () => PopQuestion(),
                    onExit: () => PopAll()
                ),
                _ => throw new NotImplementedException()
            };
        }

        // Public
        public UserDialogCoordinator()
        {
            QuestionStack = new Stack<UserQuestion>();
        }

        public void Start()
        {
            PushQuestion(createMainMenu());
            while (!ShouldEnd)
            {
                LoadTop();
            }
            EndProgram();
        }
    }
}
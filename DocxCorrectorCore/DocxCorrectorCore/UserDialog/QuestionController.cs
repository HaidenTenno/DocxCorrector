using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxCorrectorCore.UserDialog
{
    public abstract class QuestionController
    {
        // Protected
        protected QuestionsNavigationController? NavigationController;

        protected string QuestionInfo;

        protected int? GetUserAnswerInt()
        {
            string userAnswer = Console.ReadLine();
            bool result = int.TryParse(userAnswer, out int userAnserInt);

            if (!result) { return null; }
            return userAnserInt;
        }

        protected List<string> GetUserAnswerString()
        {
            string fullUserAnser = Console.ReadLine();
            List<string> userAnsers = fullUserAnser.Split(" ").ToList();
            return userAnsers;
        }

        protected void OnBack()
        {
            NavigationController?.PopQuestionController();
        }

        protected void OnExit()
        {
            NavigationController?.PopAllQuestionControllers();
        }

        // Public
        public void SetNavigationController(QuestionsNavigationController navigationController)
        {
            NavigationController = navigationController;
        }

        public QuestionController(string questionInfo)
        {
            QuestionInfo = questionInfo;
        }

        public abstract void Load();
    }

    public class IntAnswerQuestionController : QuestionController
    {
        // Protected
        protected readonly List<(string info, Action action)> Actions;

        protected List<(string info, Action action)> GetActionsToShow()
        {
            List<(string info, Action action)> actionsToShow = new List<(string info, Action action)>(Actions);
            if (NavigationController?.GetQuestionControllers().Count > 1) { actionsToShow.Add(("Назад", () => OnBack())); }
            actionsToShow.Add(("Выход", () => OnExit()));
            return actionsToShow;
        }

        // Public
        public IntAnswerQuestionController(List<(string info, Action action)> actions) : base("Выберите функцию")
        {
            Actions = actions;
        }

        public override void Load()
        {
            List<(string info, Action action)> actionsToShow = GetActionsToShow();
            Console.WriteLine(QuestionInfo);

            for (int i = 0; i < actionsToShow.Count; i++)
            {
                Console.WriteLine($"{i}: {actionsToShow[i].info}");
            }

            int? userAnser = GetUserAnswerInt();

            Console.Clear();

            if ((userAnser == null) | (userAnser >= actionsToShow.Count) | (userAnser < 0))
            {
                Console.WriteLine("Недопустимая операция");
                return;
            }

            Console.WriteLine(actionsToShow[(int)userAnser!].info);
            actionsToShow[(int)userAnser!].action();
        }
    }

    public class StringAnswerQuestionController : QuestionController
    {
        // Protected
        protected List<string> UserAnswer = new List<string>();

        protected bool CheckIfBackOrExit()
        {
            if (UserAnswer.Count == 0)
            {
                Console.WriteLine("Введите ответ");
                return false;
            }

            if (UserAnswer.Count == 1)
            {
                switch (UserAnswer.First())
                {
                    case "0":
                        OnBack();
                        return true;
                    case "1":
                        OnExit();
                        return true;
                }
            }

            return false;
        }

        protected bool CheckIfWrongArgumentsCountPassed(int requiredCount)
        {
            if (UserAnswer.Count != requiredCount)
            {
                Console.WriteLine("Некорректное число аргументов");
                return true;
            }
            return false;
        }

        // Public
        public StringAnswerQuestionController(string questionInfo) : base(questionInfo) { }

        public override void Load()
        {
            Console.WriteLine(QuestionInfo);

            Console.WriteLine("0: Назад");
            Console.WriteLine("1: Выход");

            UserAnswer = GetUserAnswerString();
            Console.Clear();
        }
    }
}


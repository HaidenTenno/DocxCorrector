using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxCorrectorCore.App
{
    public enum UserQuestionType
    {
         Print,
         PageProperties,
         SectionProperties,
         HeadersFooters,
         ParagraphProperties,
         NormalizedParagraphProperties
    }


    public abstract class UserQuestion
    {
        // Protected
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

        // Public
        public UserQuestion(string questionInfo)
        {
            QuestionInfo = questionInfo;
        }

        public abstract void Load();
    }

    public class IntAnswerQuestion : UserQuestion
    {
        // Private
        private List<(string info, Action action)> Actions;

        // Public
        public IntAnswerQuestion(List<(string info, Action action)> actions) : base("Выберите функцию")
        {
            Actions = actions;
        }

        public override void Load()
        {
            Console.WriteLine(QuestionInfo);

            for (int i = 0; i < Actions.Count; i++)
            {
                Console.WriteLine($"{i}: {Actions[i].info}");
            }

            int? userAnser = GetUserAnswerInt();

            Console.Clear();

            if ((userAnser == null) | (userAnser >= Actions.Count) | (userAnser < 0))
            {
                Console.WriteLine("Недопустимая операция");
                return;
            }

            Console.WriteLine(Actions[(int)userAnser!].info);
            Actions[(int)userAnser!].action();
        }
    }

    public class StringAnswerQuestion : UserQuestion
    {
        // Protected
        protected Action OnBack;
        protected Action OnExit;

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

        // Public
        public StringAnswerQuestion(string questionInfo, Action onBack, Action onExit) : base(questionInfo)
        {
            OnBack = onBack;
            OnExit = onExit;
        }

        public override void Load()
        {
            Console.WriteLine(QuestionInfo);

            Console.WriteLine("0: Назад");
            Console.WriteLine("1: Выход");

            UserAnswer = GetUserAnswerString();
            Console.Clear();
        }
    }

    public class PrintQuestion : StringAnswerQuestion
    {
        // Public
        public PrintQuestion(Action onBack, Action onExit) : base("Введите путь к файлу для печати", onBack, onExit) { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 1)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.PrintParagraphs(UserAnswer[0]);
        }
    }

    public class PagePropertiesJSONQuestion : StringAnswerQuestion
    {
        // Public
        public PagePropertiesJSONQuestion(Action onBack, Action onExit) : base("Введите путь к анализируемуму файлу и путь к файлу для записи свойств страниц", onBack, onExit) { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 2)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GeneratePagesPropertiesJSON(UserAnswer[0], UserAnswer[1]);
        }
    }

    public class SectionPropertiesJSONQuestion : StringAnswerQuestion
    {
        // Public
        public SectionPropertiesJSONQuestion(Action onBack, Action onExit) : base("Введите путь к анализируемуму файлу и путь к файлу для записи свойств секций", onBack, onExit) { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 2)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateSectionsPropertiesJSON(UserAnswer[0], UserAnswer[1]);
        }
    }

    public class HeadersFootersPropertiesJSONQuestion : StringAnswerQuestion
    {
        // Public
        public HeadersFootersPropertiesJSONQuestion(Action onBack, Action onExit) : base("Введите: \nТип колонтитулов (0: верхний, 1: нижний) \nПуть к анализируемуму файлу \nПуть к файлу для записи свойств секций", onBack, onExit) { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 3)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            Models.HeaderFooterType? chosenHeaderFooterType = null;

            switch (UserAnswer[0])
            {
                case "0":
                    chosenHeaderFooterType = Models.HeaderFooterType.Header;
                    break;
                case "1":
                    chosenHeaderFooterType = Models.HeaderFooterType.Footer;
                    break;
                default:
                    Console.WriteLine("Выбранна некорректная опция");
                    return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();

            featuresProvider.GenerateHeadersFootersInfoJSON((Models.HeaderFooterType)chosenHeaderFooterType, UserAnswer[1], UserAnswer[2]);
        }
    }

    public class ParagraphPropertiesCSVQuestion : StringAnswerQuestion
    {
        // Public
        public ParagraphPropertiesCSVQuestion(Action onBack, Action onExit) : base("Введите путь к корневой директории для анализа файлов в поддиректориях и название результирующего файла (свойства параграфов)", onBack, onExit) { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 2)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateCSVFiles(UserAnswer[0], UserAnswer[1]);
        }
    }

    public class NormalizedParagraphPropertiesCSVQuestion : StringAnswerQuestion
    {
        // Public
        public NormalizedParagraphPropertiesCSVQuestion(Action onBack, Action onExit) : base("Введите путь к корневой директории для анализа файлов в поддиректориях и название результирующего файла (нормализованные свойства параграфов)", onBack, onExit) { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 2)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.GenerateNormalizedCSVFiles(UserAnswer[0], UserAnswer[1]);
        }
    }
}

/*
("Генерация CSV для свойств параграфов", () => { }),
("Генерация CSV для нормализованных свойств параграфов", () => { }),
 */

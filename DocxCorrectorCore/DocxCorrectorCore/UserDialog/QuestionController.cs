using System;
using System.Collections.Generic;
using System.Linq;
using DocxCorrectorCore.App;

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

    public class PrintQuestionController : StringAnswerQuestionController
    {
        // Public
        public PrintQuestionController() : base("Введите: \nПуть к документу, параграфы которого нужно вывести в консоль") { }

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

    public class PagePropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public PagePropertiesJSONQuestionController() : base("Введите: \nПуть к документу \nПуть к директории для сохранения JSON файла со свойствами страниц") { }

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

    public class SectionPropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public SectionPropertiesJSONQuestionController() : base("Введите: \nПуть к документу \nПуть к директории для сохранения JSON файла со свойствами секций") { }

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

    public class HeadersFootersPropertiesJSONQuestionController : StringAnswerQuestionController
    {
        // Public
        public HeadersFootersPropertiesJSONQuestionController() : base("Введите: \nТип колонтитулов (0: верхний, 1: нижний) \nПуть к документу \nПуть к директории для сохранения JSON файла со свойствами колонтитулов") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 3)
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

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
                    Console.WriteLine("Выбранна некорректная опция");
                    return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();

            featuresProvider.GenerateHeadersFootersInfoJSON((Models.HeaderFooterType)chosenHeaderFooterType, UserAnswer[1], UserAnswer[2]);
        }
    }

    public class ParagraphPropertiesCSVQuestionController : StringAnswerQuestionController
    {
        // Public
        public ParagraphPropertiesCSVQuestionController() : base("Введите: \nПуть к корневой директории, в поддиректориях которой находятся документы, для которых нужно создать CSV со свойствами параграфов") { }

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
            featuresProvider.GenerateCSVFiles(UserAnswer[0]);
        }
    }

    public class NormalizedParagraphPropertiesCSVQuestionController : StringAnswerQuestionController
    {
        // Public
        public NormalizedParagraphPropertiesCSVQuestionController() : base("Введите: \nПуть к корневой директории, в поддиректориях которой находятся документы, для которых нужно создать CSV с нормализованными свойствами параграфов") { }

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
            featuresProvider.GenerateNormalizedCSVFiles(UserAnswer[0]);
        }
    }

    public class SaveDocumentAsPdfQuestionController : StringAnswerQuestionController
    {
        // Public
        public SaveDocumentAsPdfQuestionController() : base("Введите: \nПуть к docx файлу, который необходимо сохранить как pdf \nПуть к директории для сохранения результата") { }

        public override void Load()
        {
            base.Load();

            if (CheckIfBackOrExit()) { return; }

            if (UserAnswer.Count != 2 )
            {
                Console.WriteLine("Некорректное число аргументов");
                return;
            }

            FeaturesProvider featuresProvider = new FeaturesProvider();
            featuresProvider.SaveDocumentAsPdf(UserAnswer[0], UserAnswer[1]);
        }
    }

    public class SavePagesAsPdfQuestionController : StringAnswerQuestionController
    {
        // Public
        public SavePagesAsPdfQuestionController() : base("Введите: \nПуть к docx файлу, который необходимо сохранить как pdf \nПуть к директории для сохранения результата") { }

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
            featuresProvider.SavePagesAsPdf(UserAnswer[0], UserAnswer[1]);
        }
    }
}


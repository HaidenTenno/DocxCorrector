using System;
using System.Collections.Generic;

namespace DocxCorrectorCore.UserDialog
{
    public sealed class QuestionsNavigationController
    {
        // Private
        private readonly Stack<QuestionController> QuestionControllers;

        private void LoadTop()
        {
            if (QuestionControllers.Count == 0) { return; }
            QuestionControllers.Peek().Load();
        }

        private void EndProgram()
        {
            Console.WriteLine("\nEnd of program");
            Console.ReadLine();
        }

        // Public
        public QuestionsNavigationController(QuestionController? rootQuestionController = null)
        {
            QuestionControllers = new Stack<QuestionController>();
            if (rootQuestionController != null)
            {
                PushQuestionController(rootQuestionController);
            }
        }

        public Stack<QuestionController> GetQuestionControllers()
        {
            return new Stack<QuestionController>(QuestionControllers);
        }

        public void PushQuestionController(QuestionController questionController)
        {
            questionController.SetNavigationController(this);
            QuestionControllers.Push(questionController);
        }

        public void PopQuestionController()
        {
            if (QuestionControllers.Count == 0) { return; }
            QuestionControllers.Pop();
        }

        public void PopAllQuestionControllers()
        {
            QuestionControllers.Clear();
        }

        public void Run()
        {
            while (QuestionControllers.Count != 0)
            {
                LoadTop();
            }
            EndProgram();
        }
    }
}

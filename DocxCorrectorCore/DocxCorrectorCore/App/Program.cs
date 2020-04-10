using System;
using DocxCorrectorCore.UserDialog;

namespace DocxCorrectorCore.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            UserDialogCoordinator userDialogCoordinator = new UserDialogCoordinator(new QuestionsNavigationController());
            userDialogCoordinator.Start();
        }
    }
}

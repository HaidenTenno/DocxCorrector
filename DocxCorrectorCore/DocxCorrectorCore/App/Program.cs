using System;

namespace DocxCorrectorCore.App
{
    class Program
    {
        // Точка входа
        static void Main(string[] args)
        {
            UserDialogCoordinator userDialogCoordinator = new UserDialogCoordinator();
            userDialogCoordinator.Start();
        }
    }
}

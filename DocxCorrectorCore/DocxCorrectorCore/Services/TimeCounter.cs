using System;

namespace DocxCorrectorCore.Services
{
    public static class TimeCounter
    {
        // Выполняет переданный метод, и выводит время выполнения
        public static void CountTime(Action action)
        {
            var startTime = System.Diagnostics.Stopwatch.StartNew();
            action();

            startTime.Stop();

            var resultTime = startTime.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:000}",
                resultTime.Hours,
                resultTime.Minutes,
                resultTime.Seconds,
                resultTime.Milliseconds
             );
            Console.WriteLine(elapsedTime);
        }
    }
}

using System;

namespace DocxCorrectorCore.Services
{
    public static class TimeCounter
    {
        public enum ResultType
        {
            Full,
            TotalMilliseconds
        }

        // Выполняет переданный метод, и возвращает время выполнения
        public static string GetExecutionTime(Action action, ResultType resultType = ResultType.Full)
        {
            var startTime = System.Diagnostics.Stopwatch.StartNew();
            action();
            startTime.Stop();

            var resultTime = startTime.Elapsed;

            string elapsedTime = resultType switch
            {
                ResultType.Full => string.Format("{0:00}:{1:00}:{2:00}.{3:000}",
                    resultTime.Hours,
                    resultTime.Minutes,
                    resultTime.Seconds,
                    resultTime.Milliseconds
                ),
                ResultType.TotalMilliseconds => $"{Math.Round(resultTime.TotalMilliseconds)}ms",
                _ => throw new NotImplementedException(),
            };

            return elapsedTime;
        }

        // Выполняет переданный метод, и выводит время выполнения
        public static void LogExecutionTime(Action action, ResultType resultType = ResultType.Full)
        {
            string elapsedTime = GetExecutionTime(action, resultType);
            Console.WriteLine(elapsedTime);
        }
    }
}

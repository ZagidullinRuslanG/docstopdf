namespace DocxToPdfConverter.Core
{
    /// <summary>
    /// Простая реализация логгера для консоли
    /// </summary>
    public class ConsoleLogger : ILogger
    {
        public void LogInfo(string message)
        {
            Console.WriteLine($"[INFO] {message}");
        }

        public void LogWarning(string message)
        {
            Console.WriteLine($"[WARNING] {message}");
        }

        public void LogError(string message, Exception? exception = null)
        {
            Console.WriteLine($"[ERROR] {message}");
            if (exception != null)
            {
                Console.WriteLine($"Exception: {exception}");
            }
        }
    }
}

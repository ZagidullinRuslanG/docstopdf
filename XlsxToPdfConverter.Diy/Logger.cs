using System;

namespace XlsxToPdfConverter.Diy
{
    public class Logger
    {
        public void Warn(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"[WARN] {message}");
            Console.ResetColor();
        }

        public void Error(Exception ex, string message = null)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] {message}\n{ex}");
            Console.ResetColor();
        }
    }

    public static class LoggerFactory
    {
        public static Logger Get() => new Logger();
    }
} 
namespace DocxToPdfConverter.Core
{
    /// <summary>
    /// Интерфейс для логгирования
    /// </summary>
    public interface ILogger
    {
        void LogInfo(string message);
        void LogWarning(string message);
        void LogError(string message, Exception? exception = null);
    }
}

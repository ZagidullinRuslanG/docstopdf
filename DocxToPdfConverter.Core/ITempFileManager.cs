namespace DocxToPdfConverter.Core
{
    /// <summary>
    /// Интерфейс для управления временными файлами
    /// </summary>
    public interface ITempFileManager : IDisposable
    {
        /// <summary>
        /// Создает временный файл и возвращает его путь
        /// </summary>
        string CreateTempFile(string extension = ".tmp");

        /// <summary>
        /// Добавляет файл в список для автоматической очистки
        /// </summary>
        void RegisterTempFile(string filePath);

        /// <summary>
        /// Очищает все зарегистрированные временные файлы
        /// </summary>
        void Cleanup();
    }
}

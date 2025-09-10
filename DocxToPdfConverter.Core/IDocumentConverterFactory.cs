namespace DocxToPdfConverter.Core
{
    /// <summary>
    /// Интерфейс фабрики для создания конвертеров документов
    /// </summary>
    public interface IDocumentConverterFactory
    {
        /// <summary>
        /// Создает конвертер для указанного типа файла
        /// </summary>
        /// <param name="fileExtension">Расширение файла (например, ".docx", ".xlsx")</param>
        /// <param name="logger">Логгер (опционально)</param>
        /// <returns>Конвертер документов</returns>
        /// <exception cref="NotSupportedException">Если формат файла не поддерживается</exception>
        IDocumentConverter CreateConverter(string fileExtension, ILogger? logger = null);

        /// <summary>
        /// Создает конвертер на основе пути к файлу
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        /// <param name="logger">Логгер (опционально)</param>
        /// <returns>Конвертер документов</returns>
        IDocumentConverter CreateConverterFromFile(string filePath, ILogger? logger = null);

        /// <summary>
        /// Получает список поддерживаемых расширений файлов
        /// </summary>
        /// <returns>Массив поддерживаемых расширений</returns>
        string[] GetSupportedExtensions();
    }
}

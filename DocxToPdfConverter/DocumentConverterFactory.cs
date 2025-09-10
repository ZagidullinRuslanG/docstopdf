using DocxToPdfConverter.Core;
using XlsxToPdfConverter.Diy;

namespace DocxToPdfConverter
{
    /// <summary>
    /// Фабрика для создания конвертеров документов
    /// </summary>
    public class DocumentConverterFactory : IDocumentConverterFactory
    {
        /// <summary>
        /// Создает конвертер для указанного типа файла
        /// </summary>
        /// <param name="fileExtension">Расширение файла (например, ".docx", ".xlsx")</param>
        /// <param name="logger">Логгер (опционально)</param>
        /// <returns>Конвертер документов</returns>
        /// <exception cref="NotSupportedException">Если формат файла не поддерживается</exception>
        public IDocumentConverter CreateConverter(string fileExtension, Core.ILogger? logger = null)
        {
            if (string.IsNullOrEmpty(fileExtension))
                throw new ArgumentException("Расширение файла не может быть пустым", nameof(fileExtension));

            var extension = fileExtension.ToLowerInvariant();
            
            return extension switch
            {
                ".docx" => new DocxToPdfConverter(logger),
                ".xlsx" => new XlsxToPdfDiyConverter(),
                _ => throw new NotSupportedException($"Формат файла '{fileExtension}' не поддерживается")
            };
        }

        /// <summary>
        /// Создает конвертер на основе пути к файлу
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        /// <param name="logger">Логгер (опционально)</param>
        /// <returns>Конвертер документов</returns>
        public IDocumentConverter CreateConverterFromFile(string filePath, Core.ILogger? logger = null)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("Путь к файлу не может быть пустым", nameof(filePath));

            var extension = Path.GetExtension(filePath);
            return CreateConverter(extension, logger);
        }

        /// <summary>
        /// Получает список поддерживаемых расширений файлов
        /// </summary>
        /// <returns>Массив поддерживаемых расширений</returns>
        public string[] GetSupportedExtensions()
        {
            return new[] { ".docx", ".xlsx" };
        }
    }
}

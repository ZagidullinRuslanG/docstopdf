using System.IO;

namespace DocxToPdfConverter.Core
{
    /// <summary>
    /// Единый интерфейс для конвертации документов в PDF
    /// </summary>
    public interface IDocumentConverter
    {
        /// <summary>
        /// Конвертирует документ из файла в PDF
        /// </summary>
        /// <param name="inputPath">Путь к входному файлу</param>
        /// <param name="outputPath">Путь к выходному PDF файлу</param>
        void Convert(string inputPath, string outputPath);

        /// <summary>
        /// Конвертирует документ из потока в PDF
        /// </summary>
        /// <param name="inputStream">Поток с документом</param>
        /// <param name="outputPath">Путь к выходному PDF файлу</param>
        void Convert(Stream inputStream, string outputPath);

        /// <summary>
        /// Поддерживаемые расширения файлов
        /// </summary>
        string[] SupportedExtensions { get; }
    }
}

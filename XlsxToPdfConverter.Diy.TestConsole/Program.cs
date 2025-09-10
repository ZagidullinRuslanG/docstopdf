using System;
using System.IO;
using System.Linq;
using DocxToPdfConverter;
using DocxToPdfConverter.Core;
using XlsxToPdfConverter.Diy;

namespace XlsxToPdfConverter.Diy.TestConsole
{
    internal class Program
    {
        // note: Консольное приложение для профайлера.

        private static void Main(string[] args)
        {
            var logger = new DocxToPdfConverter.Core.ConsoleLogger();
            
            if (args.Length == 1 && Directory.Exists(args[0]))
            {
                ProcessDirectory(args[0], logger);
                return;
            }
            
            if (args.Length == 2 && File.Exists(args[0]))
            {
                ProcessSingleFile(args[0], args[1], logger);
                return;
            }
            
            ShowUsage();
        }

        private static void ProcessDirectory(string directoryPath, DocxToPdfConverter.Core.ILogger logger)
        {
            try
            {
                logger.LogInfo($"Обработка директории: {directoryPath}");
                string jobId = Guid.NewGuid().ToString("N");
                
                // Получаем все поддерживаемые файлы
                var factory = new DocumentConverterFactory();
                var supportedExtensions = factory.GetSupportedExtensions();
                var files = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(f => supportedExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                    .Where(f => !Path.GetFileName(f).StartsWith("~$")) // пропускаем временные файлы Word
                    .ToArray();

                foreach (string file in files)
                {
                    try
                    {
                        var converter = factory.CreateConverterFromFile(file, logger);
                        string outputPath = file + $".{jobId}.pdf";
                        converter.Convert(file, outputPath);
                        logger.LogInfo($"Конвертирован: {Path.GetFileName(file)} -> {Path.GetFileName(outputPath)}");
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Ошибка при конвертации {file}: {ex.Message}", ex);
                    }
                }
                
                logger.LogInfo($"Обработка директории завершена. Обработано файлов: {files.Length}");
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при обработке директории {directoryPath}: {ex.Message}", ex);
            }
        }

        private static void ProcessSingleFile(string inputPath, string outputPath, DocxToPdfConverter.Core.ILogger logger)
        {
            try
            {
                logger.LogInfo($"Конвертация файла: {inputPath} -> {outputPath}");
                
                var factory = new DocumentConverterFactory();
                var converter = factory.CreateConverterFromFile(inputPath, logger);
                converter.Convert(inputPath, outputPath);
                
                logger.LogInfo($"Конвертация завершена успешно");
            }
            catch (NotSupportedException ex)
            {
                logger.LogError($"Неподдерживаемый формат файла: {ex.Message}");
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при конвертации: {ex.Message}", ex);
            }
        }

        private static void ShowUsage()
        {
            Console.WriteLine("Использование:");
            Console.WriteLine("  dotnet run --project ... <file.docx|file.xlsx> <output.pdf>");
            Console.WriteLine("  dotnet run --project ... <folder>");
            Console.WriteLine();
            Console.WriteLine("Поддерживаемые форматы:");
            var factory = new DocumentConverterFactory();
            foreach (var ext in factory.GetSupportedExtensions())
            {
                Console.WriteLine($"  {ext}");
            }
        }
    }
}
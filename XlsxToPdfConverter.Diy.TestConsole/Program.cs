using System;
using System.IO;
using DocxToPdfConverter;

namespace XlsxToPdfConverter.Diy.TestConsole
{
    internal class Program
    {
        // note: Консольное приложение для профайлера.

        private static void Main(string[] args)
        {
            if (args.Length == 1 && Directory.Exists(args[0]))
            {
                string dir = args[0];
                var xlsxConverter = new XlsxToPdfDiyConverter();
                string jobId = Guid.NewGuid().ToString("N");
                foreach (string file in Directory.GetFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly))
                {
                    xlsxConverter.Convert(file, file + $".{jobId}.pdf");
                    Console.WriteLine($"XLSX to PDF: {file} -> {file}.{jobId}.pdf");
                }
                foreach (string file in Directory.GetFiles(dir, "*.docx", SearchOption.TopDirectoryOnly))
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue; // пропускаем временные файлы Word
                    DocxToPdfConverter.DocxToPdfConverter.Convert(file, file + $".{jobId}.pdf");
                    Console.WriteLine($"DOCX to PDF: {file} -> {file}.{jobId}.pdf");
                }
                return;
            }
            if (args.Length == 2 && File.Exists(args[0]))
            {
                if (args[0].EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    DocxToPdfConverter.DocxToPdfConverter.Convert(args[0], args[1]);
                    Console.WriteLine($"DOCX to PDF: {args[0]} -> {args[1]}");
                }
                else if (args[0].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    IXlsxToPdfConverter xlsxConverter = new XlsxToPdfDiyConverter();
                    xlsxConverter.Convert(args[0], args[1]);
                    Console.WriteLine($"XLSX to PDF: {args[0]} -> {args[1]}");
                }
                else
                {
                    Console.WriteLine("Unsupported file type.");
                }
                return;
            }
            Console.WriteLine("Usage:\n  dotnet run --project ... <file.docx|file.xlsx> <output.pdf>\n  dotnet run --project ... <folder>");
        }
    }
}
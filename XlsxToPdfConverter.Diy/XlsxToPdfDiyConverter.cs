using System;
using System.IO;
using DocxToPdfConverter.Core;
using NpoiHelpers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Преобразование документа XLSX в PDF.
    /// </summary>
    public class XlsxToPdfDiyConverter : IXlsxToPdfConverter, IDocumentConverter
    {
        public string[] SupportedExtensions => new[] { ".xlsx" };

        public void Convert(string xlsxPath, string pdfPath)
        {
            var mergedCellsAddresses = new MergedCellsDataXlsxDocReader().GetMergedCellsData(xlsxPath);
            IWorkbook workbook = new XSSFWorkbook(xlsxPath);
            Convert(workbook, mergedCellsAddresses, pdfPath);
            workbook.Close();
        }

        public void Convert(Stream inputStream, string pdfPath)
        {
            if (inputStream == null)
                throw new ArgumentNullException(nameof(inputStream));
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentException("Путь к выходному файлу не может быть пустым", nameof(pdfPath));

            try
            {
                // Создаем копию потока для чтения данных о объединенных ячейках
                var streamCopy = new MemoryStream();
                inputStream.CopyTo(streamCopy);
                streamCopy.Position = 0;
                
                var mergedCellsAddresses = new MergedCellsDataXlsxDocReader().GetMergedCellsData(streamCopy);
                
                // Сбрасываем позицию исходного потока
                inputStream.Position = 0;
                IWorkbook workbook = new XSSFWorkbook(inputStream);
                Convert(workbook, mergedCellsAddresses, pdfPath);
                workbook.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при конвертации из потока: {ex.Message}");
                throw;
            }
        }

        public void Convert(MemoryStream xlsxFile, string pdfPath)
        {
            Convert((Stream)xlsxFile, pdfPath);
        }

        private void Convert(
            IWorkbook workbook,
            MergedCellsData mergedCellsData,
            string pdfPath)
        {
            using (var pdfWriter = new PdfWriter())
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);

                    var sheetWriter = new XlSheetToPdfWriter(workbook, pdfWriter, sheet, i + 1, mergedCellsData);
                    sheetWriter.Write();
                }

                pdfWriter.Save(pdfPath);
                Console.WriteLine($"PDF создан: {pdfPath}");
            }
        }
    }
}

using System;
using System.IO;
using NpoiHelpers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Преобразование документа XLSX в PDF.
    /// </summary>
    public class XlsxToPdfDiyConverter : IXlsxToPdfConverter
    {
        public void Convert(string xlsxPath, string pdfPath)
        {
            var mergedCellsAddresses = new MergedCellsDataXlsxDocReader().GetMergedCellsData(xlsxPath);
            IWorkbook workbook = new XSSFWorkbook(xlsxPath);
            Convert(workbook, mergedCellsAddresses, pdfPath);
            workbook.Close();
        }

        public void Convert(MemoryStream xlsxFile, string pdfPath)
        {
            var mergedCellsAddresses = new MergedCellsDataXlsxDocReader().GetMergedCellsData(xlsxFile);
            IWorkbook workbook = new XSSFWorkbook(xlsxFile);
            Convert(workbook, mergedCellsAddresses, pdfPath);
            workbook.Close();
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

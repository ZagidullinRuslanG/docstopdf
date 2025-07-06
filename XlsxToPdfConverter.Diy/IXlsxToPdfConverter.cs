namespace XlsxToPdfConverter.Diy
{
    public interface IXlsxToPdfConverter
    {
        // Интерфейс-заглушка для совместимости
        void Convert(string xlsxPath, string pdfPath);
        void Convert(System.IO.MemoryStream xlsxFile, string pdfPath);
    }
} 
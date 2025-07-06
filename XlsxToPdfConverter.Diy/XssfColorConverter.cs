using System;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Преобразование цвета из XSSF (библиотека NPOI) в другие.
    /// </summary>
    public static class XssfColorConverter
    {
        /// <summary>
        /// Сдвиг индекса в палетте цветов NPOI.
        /// </summary>
        private const short NpoiColorIndexOffset = 8;
        
        /// <summary>
        /// Преобразовать цвет в RGB набор байт.
        /// </summary>
        public static byte[] ConvertToRgb(XSSFColor xssfColor)
        {
            if (xssfColor.RGB != null)
            {
                return xssfColor.RGB;
            }

            if (xssfColor.IsIndexed)
            {
                var palette = new HSSFPalette(new PaletteRecord());
                var index = xssfColor.Index + NpoiColorIndexOffset;
                return palette.GetColor((short)index).RGB;
            }

            throw new NotSupportedException("Формат цвета не поддерживается.");
        }
    }
}
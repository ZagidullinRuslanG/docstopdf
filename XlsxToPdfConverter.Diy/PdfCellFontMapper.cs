using System;
using XlsxToPdfConverter.Diy;
using PdfSharp.Drawing;
using NPOI.SS.UserModel;

namespace XlsxToPdfConverter.Diy
{
    /// <summary>
    /// Преборазование объекта-шрифта NPOI в объект-шрифт PDFSharp.
    /// </summary>
    internal class PdfCellFontMapper
    {
        private readonly ILogger logger = new ConsoleLogger();

        public XFont Get(IFont xlFont)
        {
            try
            {
                return new XFont(
                    xlFont.FontName,
                    xlFont.FontHeightInPoints,
                    GetFontStyle(xlFont));
            }
            catch (Exception exception)
            {
                logger.Error(exception, $"Проблемы со шрифтом {xlFont?.FontName}, {xlFont?.FontHeightInPoints}");
                throw;
            }
        }

        private XFontStyleEx GetFontStyle(IFont xlFont)
        {
            return XFontStyleEx.Regular |
                   (xlFont.IsBold ? XFontStyleEx.Bold : 0) |
                   (xlFont.IsItalic ? XFontStyleEx.Italic : 0) |
                   (xlFont.Underline != FontUnderlineType.None ? XFontStyleEx.Underline : 0) |
                   (xlFont.IsStrikeout ? XFontStyleEx.Strikeout : 0);
        }
    }
}

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NpoiHelpers
{
    public static class WorkbookExtensions
    {
        public static IFont CreateFont(this IWorkbook workbook, double fontHeightInPoints, string fontName, bool isBold)
        {
            var font = workbook.CreateFont();
            font.FontHeightInPoints = fontHeightInPoints;
            font.FontName = fontName;
            font.IsBold = isBold;

            return font;
        }
    }
}